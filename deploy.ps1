# =============================================================================
# Build & Package GenerateDeliveryReports
# =============================================================================
# Usage:
#   .\deploy.ps1 -TargetPath "\\MACHINE\C$\Apps\GenerateDeliveryReports"
#   .\deploy.ps1 -TargetPath "D:\Apps\GenerateDeliveryReports"  (local target)
#
# The script will:
#   1. Clean previous build output
#   2. Publish the app self-contained
#   3. Copy the PPTX report template into the publish output
#   4. Assemble CommonFiles (xlsx data files + PPTX template)
#   5. Patch appsettings.json with target-machine paths
#   6. Stage the GenerateSprintDashboard folder into the package
#   7. Robocopy the entire package to TargetPath
#
# Output: a ready-to-copy package under .\package\
#   package\
#     Web\                     <- published .NET app (appsettings.json already patched)
#     CommonFiles\
#       Templates\             <- PPTX report template
#       Files\                 <- xlsx data files mirrored from OneDrive
#     GenerateSprintDashboard\ <- Python sprint dashboard script + template
#
# TargetPath is required so appsettings.json paths are baked in correctly
# for the machine the package will be deployed to.
#
# The app binds to http://*:5158 (configured in Program.cs).
# =============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,

    [string]$Configuration = "Release",
    [string]$Runtime       = "win-x64"
)

$ErrorActionPreference = "Stop"
$ProjectPath         = Join-Path $PSScriptRoot "GenerateDeliveryReports\GenerateDeliveryReports.csproj"
$SrcAppSettings      = Join-Path $PSScriptRoot "GenerateDeliveryReports\appsettings.json"
$PublishDir          = Join-Path $PSScriptRoot "publish"
$TemplateSrc         = Join-Path $PSScriptRoot "GenerateDeliveryReports.Data\Templates"
$PackageDir          = Join-Path $PSScriptRoot "package"
$PackageWeb          = Join-Path $PSScriptRoot "package\Web"
$PackageCommon       = Join-Path $PSScriptRoot "package\CommonFiles"
$SprintDashboardSrc  = Join-Path $PSScriptRoot "GenerateSprintDashboard"
$PackageSprintDash   = Join-Path $PSScriptRoot "package\GenerateSprintDashboard"

# Locate dotnet.exe
$dotnetCmd = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnetCmd) {
    $defaultPaths = @(
        "$env:ProgramFiles\dotnet\dotnet.exe",
        "${env:ProgramFiles(x86)}\dotnet\dotnet.exe",
        "$env:LOCALAPPDATA\Microsoft\dotnet\dotnet.exe"
    )
    foreach ($p in $defaultPaths) {
        if (Test-Path $p) { $dotnetCmd = $p; break }
    }
    if (-not $dotnetCmd) {
        Write-Host "ERROR: 'dotnet' not found. Install .NET SDK or add it to PATH." -ForegroundColor Red
        exit 1
    }
}
$dotnet = if ($dotnetCmd -is [string]) { $dotnetCmd } else { $dotnetCmd.Source }
Write-Host "Using dotnet: $dotnet" -ForegroundColor Gray

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " GenerateDeliveryReports - Package Script" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# Step 1: Clean previous publish output and package staging
Write-Host "`n[1/7] Cleaning previous build output..." -ForegroundColor Yellow
foreach ($dir in @($PublishDir, $PackageDir)) {
    if (Test-Path $dir) {
        Remove-Item -Recurse -Force $dir
        Write-Host "  Removed $dir" -ForegroundColor Gray
    }
}

# Step 2: Publish self-contained
Write-Host "`n[2/7] Publishing ($Configuration | $Runtime | self-contained)..." -ForegroundColor Yellow
& $dotnet publish $ProjectPath -c $Configuration -r $Runtime --self-contained -o $PublishDir
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Publish failed." -ForegroundColor Red
    exit 1
}
Write-Host "Publish succeeded." -ForegroundColor Green

# Step 3: Copy PPTX template into publish output
Write-Host "`n[3/7] Copying report template into publish output..." -ForegroundColor Yellow
$TemplateDestDir = Join-Path $PublishDir "Templates"
if (-not (Test-Path $TemplateDestDir)) {
    New-Item -ItemType Directory -Path $TemplateDestDir | Out-Null
}
Copy-Item -Path (Join-Path $TemplateSrc "*") -Destination $TemplateDestDir -Force
Write-Host "Template copied to $TemplateDestDir" -ForegroundColor Green

# Step 4: Assemble CommonFiles -- mirror xlsx files from OneDrive and copy PPTX template
Write-Host "`n[4/7] Assembling CommonFiles..." -ForegroundColor Yellow

# Read OneDriveLocation from the source appsettings.json on this machine.
# ConvertFrom-Json is safe here -- we are only reading, not writing back.
$srcJson     = Get-Content $SrcAppSettings -Raw | ConvertFrom-Json
$oneDriveSrc = $srcJson.AppSettings.OneDriveLocation
if (-not $oneDriveSrc) {
    Write-Host "ERROR: OneDriveLocation is not set in $SrcAppSettings" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $oneDriveSrc)) {
    Write-Host "ERROR: OneDriveLocation path not found: $oneDriveSrc" -ForegroundColor Red
    exit 1
}
Write-Host "  Source OneDriveLocation: $oneDriveSrc" -ForegroundColor Gray

# Copy only the specific xlsx files referenced in appsettings.json.
# Uses robocopy per file via Start-Process to handle paths longer than 260 characters.
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

$filesDestDir  = Join-Path $PackageCommon "Files"
New-Item -ItemType Directory -Path $filesDestDir -Force | Out-Null

$oneDriveBase  = $oneDriveSrc.TrimEnd('\')
$metricsFolder = "$($srcJson.AppSettings.MetricsFolder)".TrimStart('\').TrimStart('/')
$dataFolder    = "$($srcJson.AppSettings.ReportAndDataFolder)".TrimStart('\').TrimStart('/')
$csatFolder    = "$($srcJson.AppSettings.CSAT.CSATFolder)".TrimStart('\').TrimStart('/')

Write-Host "  Staging folder  : $filesDestDir" -ForegroundColor Gray
Write-Host "  OneDrive base   : $oneDriveBase" -ForegroundColor Gray
Write-Host "  Metrics folder  : $metricsFolder" -ForegroundColor Gray
Write-Host "  Data folder     : $dataFolder" -ForegroundColor Gray
Write-Host "  CSAT folder     : $csatFolder" -ForegroundColor Gray

$xlsxCount  = 0
$xlsxFailed = 0

# Projects: MetricsSheetPath (array per project) and DataFileName
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

# CSAT: ClientSurveyFilePath -- deduplicated as multiple clients share the same survey file
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
Write-Host "  Staged $xlsxCount xlsx file(s) to CommonFiles\Files\$skippedMsg" -ForegroundColor Green

# Verify what actually landed in the staging folder
$stagedFiles = Get-ChildItem -Path $filesDestDir -Recurse -File -ErrorAction SilentlyContinue
if ($stagedFiles.Count -gt 0) {
    Write-Host "  Verified $($stagedFiles.Count) file(s) in staging folder:" -ForegroundColor Green
    $stagedFiles | ForEach-Object { Write-Host "    $($_.FullName)" -ForegroundColor DarkGray }
} else {
    Write-Host "  WARNING: Staging folder is empty -- all source paths above showed as INFO (not found)." -ForegroundColor Yellow
    Write-Host "           Check that OneDriveLocation and folder paths in appsettings.json are correct." -ForegroundColor Yellow
}

# Copy PPTX template to CommonFiles\Templates\
$commonTemplateDir = Join-Path $PackageCommon "Templates"
New-Item -ItemType Directory -Path $commonTemplateDir -Force | Out-Null
Copy-Item -Path (Join-Path $TemplateSrc "*") -Destination $commonTemplateDir -Force
Write-Host "  PPTX template copied to CommonFiles\Templates\" -ForegroundColor Green

# Step 5: Patch appsettings.json with target-machine paths and move it into package\Web\
Write-Host "`n[5/7] Patching appsettings.json with target paths..." -ForegroundColor Yellow

# Use targeted string replacement (not ConvertFrom-Json → ConvertTo-Json) to avoid
# PowerShell 5.1 silently dropping deeply-nested structures (e.g. the CSAT Clients array).
$AppSettingsPath = Join-Path $PublishDir "appsettings.json"
$content = Get-Content $AppSettingsPath -Raw

$targetOneDrive        = (Join-Path $TargetPath "CommonFiles\Files")   -replace '\\', '\\'
$targetCommon          = (Join-Path $TargetPath "CommonFiles")           -replace '\\', '\\'
$targetTemplate        = (Join-Path $TargetPath "CommonFiles\Templates\GlobalPayments-DeliveryQualitySummaryReport_Template.pptx") -replace '\\', '\\'
$targetDashboardHtml   = (Join-Path $TargetPath "GenerateSprintDashboard\sprint_dashboard.html") -replace '\\', '\\'
$targetDashboardScript = (Join-Path $TargetPath "GenerateSprintDashboard\sprint_dashboard.py")   -replace '\\', '\\'

$content = $content -replace '(?<="OneDriveLocation"\s*:\s*")[^"]*(?=")',                $targetOneDrive
$content = $content -replace '(?<="CommonFolderPath"\s*:\s*")[^"]*(?=")',                $targetCommon
$content = $content -replace '(?<="SprintMetricsReportTemplatePath"\s*:\s*")[^"]*(?=")', $targetTemplate
# Clear the hardcoded dev-machine path so the app uses its built-in fallback (wwwroot/worker-summary.html)
$content = $content -replace '(?<="WorkerSummaryFilePath"\s*:\s*")[^"]*(?=")',           ''
$content = $content -replace '(?<="SprintDashboardHtmlPath"\s*:\s*")[^"]*(?=")',         $targetDashboardHtml
$content = $content -replace '(?<="SprintDashboardScriptPath"\s*:\s*")[^"]*(?=")',       $targetDashboardScript

[System.IO.File]::WriteAllText($AppSettingsPath, $content, [System.Text.Encoding]::UTF8)

# Verify CSAT section survived the regex replacements
if ($content -notmatch '"CSAT"') {
    Write-Host "ERROR: CSAT section is missing from appsettings.json after patching. Aborting." -ForegroundColor Red
    exit 1
}
if ($content -notmatch '"Clients"') {
    Write-Host "ERROR: CSAT Clients array is missing from appsettings.json after patching. Aborting." -ForegroundColor Red
    exit 1
}
Write-Host "appsettings.json patched." -ForegroundColor Green

# Move published output into package\Web\
New-Item -ItemType Directory -Path $PackageWeb -Force | Out-Null
Copy-Item -Path (Join-Path $PublishDir "*") -Destination $PackageWeb -Recurse -Force
Write-Host "Published app staged to package\Web\" -ForegroundColor Green

# Step 6: Stage GenerateSprintDashboard folder into the package
Write-Host "`n[6/7] Staging GenerateSprintDashboard..." -ForegroundColor Yellow
if (Test-Path $SprintDashboardSrc) {
    New-Item -ItemType Directory -Path $PackageSprintDash -Force | Out-Null
    Copy-Item -Path (Join-Path $SprintDashboardSrc "*") -Destination $PackageSprintDash -Recurse -Force
    Write-Host "GenerateSprintDashboard staged to package\GenerateSprintDashboard\" -ForegroundColor Green
} else {
    Write-Host "WARNING: GenerateSprintDashboard source not found at $SprintDashboardSrc -- skipping." -ForegroundColor Yellow
}

# Step 7: Copy package to target
Write-Host "`n[7/7] Deploying package to $TargetPath ..." -ForegroundColor Yellow
if (-not (Test-Path $TargetPath)) {
    New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null
}
$proc = Start-Process "robocopy.exe" `
    -ArgumentList "`"$($PackageDir.TrimEnd('\'))`" `"$($TargetPath.TrimEnd('\'))`" /E /R:2 /W:1" `
    -Wait -PassThru -NoNewWindow
if ($proc.ExitCode -ge 8) {
    Write-Host "ERROR: Deploy failed (robocopy exit $($proc.ExitCode))." -ForegroundColor Red
    exit 1
}
Write-Host "Package deployed to $TargetPath" -ForegroundColor Green

# Summary
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host " Deploy Complete!" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Package location: $PackageDir" -ForegroundColor Yellow
Write-Host ""
Write-Host "Package layout:" -ForegroundColor Yellow
Write-Host "  package\" -ForegroundColor White
Write-Host "    Web\                          <- web application (appsettings.json pre-patched)" -ForegroundColor White
Write-Host "    CommonFiles\" -ForegroundColor White
Write-Host "      Templates\                  <- PPTX report template" -ForegroundColor White
Write-Host "      Files\                      <- xlsx data files ($xlsxCount file(s) from OneDrive)" -ForegroundColor White
Write-Host "    GenerateSprintDashboard\      <- sprint dashboard script + template" -ForegroundColor White
Write-Host ""
Write-Host "Paths baked into appsettings.json target: $TargetPath" -ForegroundColor Gray
Write-Host ""

# Show actual file counts per package subfolder
Write-Host "Package contents:" -ForegroundColor Yellow
$subFolders = @("Web", "CommonFiles\Files", "CommonFiles\Templates", "GenerateSprintDashboard")
foreach ($sub in $subFolders) {
    $subPath = Join-Path $PackageDir $sub
    if (Test-Path $subPath) {
        $count = (Get-ChildItem -Path $subPath -Recurse -File -ErrorAction SilentlyContinue).Count
        Write-Host ("  {0,-40} {1} file(s)" -f "$sub\", $count) -ForegroundColor White
    } else {
        Write-Host "  $sub\  <missing>" -ForegroundColor Yellow
    }
}
Write-Host ""
Write-Host ""
