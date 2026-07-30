# =============================================================================
# Build & Package GenerateDeliveryReports for Azure App Service
# =============================================================================
# Usage:
#   .\deploy-appservice.ps1
#   .\deploy-appservice.ps1 -Zip
#   .\deploy-appservice.ps1 -OutputPath "D:\Packages\GenerateDeliveryReports" -Zip
#
# The script will:
#   1. Clean previous package output
#   2. Publish framework-dependent (App Service supplies the .NET runtime --
#      no need to ship a self-contained win-x64 runtime like the on-prem deploy)
#   3. Bundle xlsx data files, CSAT survey files, and the PPTX template INTO the
#      package (nested under CommonFiles\, not alongside it -- see below)
#   4. Patch appsettings.json with relative paths and disable local-only features
#   5. Optionally zip the result for az webapp deploy / a GitHub Action
#
# Why this differs from deploy.ps1 (the on-prem/Task Scheduler deploy script):
#   - App Service has no OneDrive sync and no fixed local disk path to bake in,
#     so OneDriveLocation/SprintMetricsReportTemplatePath/BriefingArchiveFolder
#     are set to RELATIVE paths, resolved by AppSettings' Resolved* properties
#     against the app's own base directory (see GenerateDeliveryReports.Models).
#   - The data files must live INSIDE the published app folder for that relative
#     resolution to find them -- deploy.ps1 puts CommonFiles next to Web\, this
#     script puts it inside the package itself.
#   - No Python on App Service, so the Sprint Dashboard script path is cleared;
#     that feature already degrades gracefully when unconfigured.
#   - Framework-dependent publish (smaller, matches the platform's .NET stack)
#     instead of self-contained win-x64.
#
# This script does NOT deploy anything to Azure -- it only produces the package.
# It must run on a machine with the real OneDrive-synced files available (reads
# OneDriveLocation from appsettings.json to find them).
# =============================================================================

param(
    [string]$Configuration = "Release",
    [string]$OutputPath = "",
    [switch]$Zip
)

$ErrorActionPreference = "Stop"
$ProjectPath    = Join-Path $PSScriptRoot "GenerateDeliveryReports\GenerateDeliveryReports.csproj"
$SrcAppSettings = Join-Path $PSScriptRoot "GenerateDeliveryReports\appsettings.json"
$TemplateSrc    = Join-Path $PSScriptRoot "GenerateDeliveryReports.Data\Templates"
$PackageDir     = if ($OutputPath) { $OutputPath } else { Join-Path $PSScriptRoot "package-appservice" }

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " GenerateDeliveryReports - App Service Package" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Note: a running GenerateDeliveryReports instance only blocks this script if it's actually
# locking files under $PackageDir (e.g. your own interactive dev session publishing to this same
# path). A separately-deployed instance (Windows Service/Scheduled Task, different folder) won't
# conflict -- if dotnet publish does hit a real lock, it fails with a clear "file in use" error.

# Step 1: Clean previous package
# -Confirm:$false on top of -Force: on a non-interactive CI agent, Remove-Item throws instead of
# prompting if anything (e.g. a leftover read-only/locked file from a prior run at a persistent
# -OutputPath) would otherwise make it ask "are you sure?".
Write-Host "`n[1/5] Cleaning previous package..." -ForegroundColor Yellow
if (Test-Path $PackageDir) {
    Remove-Item -Recurse -Force -Confirm:$false $PackageDir
    Write-Host "  Removed $PackageDir" -ForegroundColor Gray
}

# Step 2: Publish framework-dependent
Write-Host "`n[2/5] Publishing ($Configuration, framework-dependent)..." -ForegroundColor Yellow
& dotnet publish $ProjectPath -c $Configuration -o $PackageDir
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Publish failed." -ForegroundColor Red
    exit 1
}
Write-Host "Publish succeeded." -ForegroundColor Green

# Step 3: Bundle data files into the package (nested inside, not alongside)
Write-Host "`n[3/5] Bundling data files into the package..." -ForegroundColor Yellow

$srcJson = Get-Content $SrcAppSettings -Raw | ConvertFrom-Json
$oneDriveBase = "$($srcJson.AppSettings.OneDriveLocation)".TrimEnd('\')
if (-not $oneDriveBase) {
    Write-Host "ERROR: OneDriveLocation is not set in $SrcAppSettings" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $oneDriveBase)) {
    Write-Host "ERROR: OneDriveLocation path not found: $oneDriveBase" -ForegroundColor Red
    Write-Host "       Run this script on a machine with the real OneDrive-synced files available." -ForegroundColor Red
    exit 1
}

$metricsFolder = "$($srcJson.AppSettings.MetricsFolder)".TrimStart('\').TrimStart('/')
$dataFolder    = "$($srcJson.AppSettings.ReportAndDataFolder)".TrimStart('\').TrimStart('/')
$csatFolder    = "$($srcJson.AppSettings.CSAT.CSATFolder)".TrimStart('\').TrimStart('/')
$filesDestDir  = Join-Path $PackageDir "CommonFiles\Files"
New-Item -ItemType Directory -Path $filesDestDir -Force | Out-Null

Write-Host "  OneDrive base : $oneDriveBase" -ForegroundColor Gray
Write-Host "  Metrics folder: $metricsFolder" -ForegroundColor Gray
Write-Host "  Data folder   : $dataFolder" -ForegroundColor Gray
Write-Host "  CSAT folder   : $csatFolder" -ForegroundColor Gray

# Copies one file via robocopy (handles paths longer than 260 characters).
function Copy-XlsxFile([string]$srcFile, [string]$dstFile) {
    $srcDir  = [System.IO.Path]::GetDirectoryName($srcFile).TrimEnd('\')
    $srcName = [System.IO.Path]::GetFileName($srcFile)
    $dstDir  = [System.IO.Path]::GetDirectoryName($dstFile).TrimEnd('\')
    if (-not [System.IO.Directory]::Exists($dstDir)) {
        [System.IO.Directory]::CreateDirectory($dstDir) | Out-Null
    }
    $tmpLog = [System.IO.Path]::GetTempFileName()
    $proc   = Start-Process "robocopy.exe" `
        -ArgumentList "`"$srcDir`" `"$dstDir`" `"$srcName`" /R:2 /W:1 /IS" `
        -Wait -PassThru -NoNewWindow -RedirectStandardOutput $tmpLog
    Remove-Item $tmpLog -ErrorAction SilentlyContinue
    return $proc.ExitCode
}

$xlsxCount  = 0
$xlsxFailed = 0

Write-Host "  Copying project metrics and data files..." -ForegroundColor Gray
foreach ($project in $srcJson.AppSettings.Projects) {
    foreach ($metricsPath in $project.MetricsSheetPath) {
        if (-not $metricsPath) { continue }
        $rel     = Join-Path $metricsFolder $metricsPath
        $srcFull = Join-Path $oneDriveBase $rel
        $code    = Copy-XlsxFile $srcFull (Join-Path $filesDestDir $rel)
        if     ($code -eq 0)  { Write-Host "    INFO: File not found -- verify or remove entry:`n          $srcFull" -ForegroundColor Cyan }
        elseif ($code -lt 8)  { $xlsxCount++; Write-Host "    OK: $srcFull" -ForegroundColor DarkGray }
        else                  { $xlsxFailed++; Write-Host "    WARNING: Could not copy $srcFull" -ForegroundColor Yellow }
    }
    if (-not $project.DataFileName) { continue }
    $rel     = Join-Path $dataFolder $project.DataFileName
    $srcFull = Join-Path $oneDriveBase $rel
    $code    = Copy-XlsxFile $srcFull (Join-Path $filesDestDir $rel)
    if     ($code -eq 0)  { Write-Host "    INFO: File not found -- verify or remove entry:`n          $srcFull" -ForegroundColor Cyan }
    elseif ($code -lt 8)  { $xlsxCount++; Write-Host "    OK: $srcFull" -ForegroundColor DarkGray }
    else                  { $xlsxFailed++; Write-Host "    WARNING: Could not copy $srcFull" -ForegroundColor Yellow }
}

$csatClientCount = @($srcJson.AppSettings.CSAT.Clients).Count
Write-Host "  Copying CSAT survey files (folder: $csatFolder | $csatClientCount client entries)..." -ForegroundColor Gray
$csatSeen = @{}
foreach ($client in $srcJson.AppSettings.CSAT.Clients) {
    $fileName = "$($client.ClientSurveyFilePath)".Trim()
    if (-not $fileName -or $csatSeen.ContainsKey($fileName)) { continue }
    $csatSeen[$fileName] = $true
    $rel     = Join-Path $csatFolder $fileName
    $srcFull = Join-Path $oneDriveBase $rel
    $code    = Copy-XlsxFile $srcFull (Join-Path $filesDestDir $rel)
    if     ($code -eq 0)  { Write-Host "    INFO: File not found -- verify or remove entry:`n          $srcFull" -ForegroundColor Cyan }
    elseif ($code -lt 8)  { $xlsxCount++; Write-Host "    OK: $srcFull" -ForegroundColor DarkGray }
    else                  { $xlsxFailed++; Write-Host "    WARNING: Could not copy $srcFull" -ForegroundColor Yellow }
}

$skippedMsg = if ($xlsxFailed -gt 0) { " ($xlsxFailed could not be copied)" } else { "" }
Write-Host "  Bundled $xlsxCount xlsx file(s)$skippedMsg" -ForegroundColor Green

$templatesDestDir = Join-Path $PackageDir "CommonFiles\Templates"
New-Item -ItemType Directory -Path $templatesDestDir -Force | Out-Null
Copy-Item -Path (Join-Path $TemplateSrc "*") -Destination $templatesDestDir -Force
$templateFile = Get-ChildItem $TemplateSrc -Filter "*.pptx" | Select-Object -First 1
if (-not $templateFile) {
    Write-Host "ERROR: No .pptx template found in $TemplateSrc" -ForegroundColor Red
    exit 1
}
Write-Host "  PPTX template bundled: $($templateFile.Name)" -ForegroundColor Green

# Step 4: Patch appsettings.json -- relative paths only; no dev-machine absolute paths;
# no Python-dependent features (App Service has no Python runtime).
Write-Host "`n[4/5] Patching appsettings.json for App Service..." -ForegroundColor Yellow

$AppSettingsPath = Join-Path $PackageDir "appsettings.json"
$content = Get-Content $AppSettingsPath -Raw

$content = $content -replace '(?<="OneDriveLocation"\s*:\s*")[^"]*(?=")',                'CommonFiles\\Files'
$content = $content -replace '(?<="SprintMetricsReportTemplatePath"\s*:\s*")[^"]*(?=")', "CommonFiles\\Templates\\$($templateFile.Name)"
$content = $content -replace '(?<="CommonFolderPath"\s*:\s*")[^"]*(?=")',                ''
$content = $content -replace '(?<="BriefingArchiveFolder"\s*:\s*")[^"]*(?=")',           'BriefingArchive'
# Python isn't available on App Service -- these already degrade gracefully when blank.
$content = $content -replace '(?<="WorkerSummaryFilePath"\s*:\s*")[^"]*(?=")',           ''
$content = $content -replace '(?<="SprintDashboardHtmlPath"\s*:\s*")[^"]*(?=")',         ''
$content = $content -replace '(?<="SprintDashboardScriptPath"\s*:\s*")[^"]*(?=")',       ''

[System.IO.File]::WriteAllText($AppSettingsPath, $content, [System.Text.Encoding]::UTF8)

if ($content -notmatch '"CSAT"') {
    Write-Host "ERROR: CSAT section is missing from appsettings.json after patching. Aborting." -ForegroundColor Red
    exit 1
}
if ($content -notmatch '"Clients"') {
    Write-Host "ERROR: CSAT Clients array is missing from appsettings.json after patching. Aborting." -ForegroundColor Red
    exit 1
}
Write-Host "appsettings.json patched." -ForegroundColor Green

# Step 5: Optionally zip for az webapp deploy / a GitHub Action
if ($Zip) {
    Write-Host "`n[5/5] Creating zip package..." -ForegroundColor Yellow
    $zipPath = "$PackageDir.zip"
    if (Test-Path $zipPath) {
        # .NET delete instead of Remove-Item: never prompts, regardless of ambient confirm
        # preference or file attributes -- Remove-Item can still try to confirm even with
        # -Force in some cases, which throws instead of prompting on a non-interactive CI agent.
        try {
            [System.IO.File]::SetAttributes($zipPath, [System.IO.FileAttributes]::Normal)
            [System.IO.File]::Delete($zipPath)
        } catch {
            Write-Host "ERROR: Could not remove existing zip at $zipPath -- $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "       It may still be open in another process (e.g. a prior deploy step)." -ForegroundColor Red
            exit 1
        }
    }
    Compress-Archive -Path (Join-Path $PackageDir "*") -DestinationPath $zipPath
    Write-Host "Created $zipPath" -ForegroundColor Green
} else {
    Write-Host "`n[5/5] Skipping zip (pass -Zip to also create one)" -ForegroundColor Yellow
}

# Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host " Package Complete!" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Package location: $PackageDir" -ForegroundColor Yellow
Write-Host ""
Write-Host "Still needed before this actually works on App Service (not handled by this script):" -ForegroundColor Yellow
Write-Host "  - Set AzureAd:TenantId / AzureAd:ClientId as App Service Configuration settings" -ForegroundColor White
Write-Host "    (or leave unset to keep Entra sign-in off, same as locally)" -ForegroundColor White
Write-Host "  - Set AzureAd__ClientSecret as an App Service Application Setting -- never commit it" -ForegroundColor White
Write-Host "  - Set AppSettings__AnthropicApiKey as an App Service Application Setting" -ForegroundColor White
Write-Host "  - Add the deployed https://<app-name>.azurewebsites.net/signin-oidc redirect URI" -ForegroundColor White
Write-Host "    to the Entra app registration once Entra is usable again" -ForegroundColor White
Write-Host "  - Local sign-in (username/password + TOTP) works with zero extra config -- the" -ForegroundColor White
Write-Host "    SQLite identity DB is created automatically under the app's own folder" -ForegroundColor White
Write-Host ""
if ($Zip) {
    Write-Host "Deploy with the Azure CLI, e.g.:" -ForegroundColor Yellow
    Write-Host "  az webapp deploy --resource-group <rg> --name <app-name> --src-path `"$PackageDir.zip`" --type zip" -ForegroundColor White
} else {
    Write-Host "Re-run with -Zip to produce a zip ready for 'az webapp deploy' or a GitHub Action." -ForegroundColor Yellow
}
Write-Host ""
