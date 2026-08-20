# Auto-Sync CommonFiles to Azure App Service
# Usage: .\auto-syncfilestoazure.ps1 -AzureAppName "app" -AzureUsername "user@email.com" -AzurePassword "pwd" -AppSettingsPath "path" -DeploymentFolder "path"

param(
    [string]$AzureHostName = "",
    [string]$AzureUsername = "",
    [string]$AzurePassword = "",
    [string]$AppSettingsPath = "",
    [string]$DeploymentFolder = ""
)
#find a way to avoid hardcoding of the Generatescript path
$GenerateScript = "C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports\generate-sprint-reports-availability.ps1"

$ErrorActionPreference = "Stop"

if (-not $AzureHostName -or -not $AzureUsername -or -not $AzurePassword) {
    Write-Host "ERROR: Missing Azure credentials" -ForegroundColor Red
    exit 1
}

Write-Host "Reading appsettings..." -ForegroundColor Yellow
$appConfig = Get-Content $AppSettingsPath -Raw | ConvertFrom-Json
$oneDriveBase = $appConfig.AppSettings.OneDriveLocation.TrimEnd('\')

# Create a temporary wrapper folder
$TempZipBase = Join-Path $env:TEMP "ZipTemp_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$TempCommonFiles = Join-Path $TempZipBase "CommonFiles"
$FilesDestDir = Join-Path $TempCommonFiles "Files"
$TemplatesDestDir = Join-Path $TempCommonFiles "Templates"

Write-Host "Preparing files..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path $FilesDestDir -Force | Out-Null
New-Item -ItemType Directory -Path $TemplatesDestDir -Force | Out-Null

# Copy XLSX files
$metricsFolder = $appConfig.AppSettings.MetricsFolder.TrimStart('\').TrimStart('/')
$dataFolder = $appConfig.AppSettings.ReportAndDataFolder.TrimStart('\').TrimStart('/')

foreach ($project in $appConfig.AppSettings.Projects) {
    foreach ($metricsPath in $project.MetricsSheetPath) {
        if (-not $metricsPath) { continue }
        $srcFull = Join-Path $oneDriveBase "$metricsFolder\$metricsPath"
        $dstFull = Join-Path $FilesDestDir "$metricsFolder\$metricsPath"

        if (Test-Path $srcFull) {
            $dstDir = Split-Path -Parent $dstFull
            New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
            Copy-Item -Path $srcFull -Destination $dstFull -Force
        }
    }

    if ($project.DataFileName) {
        $srcFull = Join-Path $oneDriveBase "$dataFolder\$($project.DataFileName)"
        $dstFull = Join-Path $FilesDestDir "$dataFolder\$($project.DataFileName)"

        if (Test-Path $srcFull) {
            $dstDir = Split-Path -Parent $dstFull
            New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
            Copy-Item -Path $srcFull -Destination $dstFull -Force
        }
    }
}

# Copy CSAT files
$csatFolder = $appConfig.AppSettings.CSAT.CSATFolder.TrimStart('\').TrimStart('/')
$csatSeen = @{}

foreach ($client in $appConfig.AppSettings.CSAT.Clients) {
    $fileName = $client.ClientSurveyFilePath.Trim()
    if (-not $fileName -or $csatSeen.ContainsKey($fileName)) { continue }
    $csatSeen[$fileName] = $true

    $srcFull = Join-Path $oneDriveBase "$csatFolder\$fileName"
    $dstFull = Join-Path $FilesDestDir "$csatFolder\$fileName"

    if (Test-Path $srcFull) {
        $dstDir = Split-Path -Parent $dstFull
        New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
        Copy-Item -Path $srcFull -Destination $dstFull -Force
    }
}

# Copy template
$templatePath = $appConfig.AppSettings.SprintMetricsReportTemplatePath
if ($templatePath -and (Test-Path $templatePath)) {
    $templateDir = Split-Path -Parent $templatePath
    if (Test-Path $templateDir) {
        Copy-Item -Path (Join-Path $templateDir "*") -Destination $TemplatesDestDir -Force -Recurse
    }
}

# Copy worker history files
$WorkerHistoryDir = Join-Path $TempCommonFiles "WorkerHistory"
if ($DeploymentFolder -and (Test-Path $DeploymentFolder)) {
    New-Item -ItemType Directory -Path $WorkerHistoryDir -Force | Out-Null

    $htmlFile = Get-ChildItem $DeploymentFolder -Filter "run-workerhistory.html" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($htmlFile) {
        Copy-Item -Path $htmlFile.FullName -Destination $WorkerHistoryDir -Force
    }

    $jsonFile = Get-ChildItem $DeploymentFolder -Filter "run-workerhistory.json" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($jsonFile) {
        Copy-Item -Path $jsonFile.FullName -Destination $WorkerHistoryDir -Force
    }
}

# Generate Sprint Reports Availability JSON
Write-Host "Generating Sprint Reports Availability JSON..." -ForegroundColor Yellow
$jsonOutputPath = Join-Path $FilesDestDir "SprintReportsAvailability.json"

if (Test-Path $GenerateScript) {
    try {
        & $GenerateScript -AppSettingsPath $AppSettingsPath -OutputJson $jsonOutputPath
        Write-Host "  OK: JSON generated" -ForegroundColor Green
        
        # VERIFY THE FILE
        if (Test-Path $jsonOutputPath) {
            $fileInfo = Get-Item $jsonOutputPath
            Write-Host "  File timestamp: $($fileInfo.LastWriteTime)" -ForegroundColor Green
            Write-Host "  File size: $($fileInfo.Length) bytes" -ForegroundColor Green
        } else {
            Write-Host "  ERROR: Generated file not found at: $jsonOutputPath" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "  WARNING: JSON generation failed: $_" -ForegroundColor Yellow
    }
}



# Remove the old code that tried to copy the JSON
# (delete the section that checks $sprintAvailJsonPath)

# Create zip
Write-Host "Creating zip..." -ForegroundColor Yellow
$ZipPath = Join-Path $env:TEMP "CommonFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
Compress-Archive -Path $TempCommonFiles -DestinationPath $ZipPath -Force

# VERIFY ZIP CONTENTS
Write-Host "`nVerifying zip contents..." -ForegroundColor Yellow
$zipContents = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
$jsonFileInZip = $zipContents.Entries | Where-Object { $_.Name -eq "SprintReportsAvailability.json" }

if ($jsonFileInZip) {
    Write-Host "  Found SprintReportsAvailability.json in zip" -ForegroundColor Green
    Write-Host "  File size in zip: $($jsonFileInZip.Length) bytes" -ForegroundColor Green
} else {
    Write-Host "  ERROR: SprintReportsAvailability.json NOT in zip!" -ForegroundColor Red
    Write-Host "  Zip contains: $($zipContents.Entries.Name -join ', ')" -ForegroundColor Yellow
}
$zipContents.Dispose()

# Upload to Kudu
Write-Host "Uploading to Azure..." -ForegroundColor Yellow
# Delete old JSON file first (to force refresh)
Write-Host "Cleaning up old files in Azure..." -ForegroundColor Yellow
$DeleteUrl = "https://$AzureHostname/api/vfs/site/wwwroot/CommonFiles/Files/SprintReportsAvailability.json"
try {
    $DeleteResponse = Invoke-WebRequest -Uri $DeleteUrl -Headers $Headers -Method DELETE -UseBasicParsing -ErrorAction SilentlyContinue
    Write-Host "  Old file deleted" -ForegroundColor Green
}
catch {
    Write-Host "  Note: Could not delete old file - will be overwritten" -ForegroundColor Yellow
}

# Upload to Kudu
Write-Host "Uploading to Azure..." -ForegroundColor Yellow
$KuduUrl = "https://$AzureHostname/api/zip/site/wwwroot/CommonFiles"
$AuthString = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($AzureUsername):$($AzurePassword)"))
$Headers = @{ "Authorization" = "Basic $AuthString"; "Content-Type" = "application/zip" }

$Response = Invoke-WebRequest -Uri $KuduUrl -Method POST -Headers $Headers -InFile $ZipPath -UseBasicParsing -TimeoutSec 300
Write-Host "Upload Status: $($Response.StatusCode)" -ForegroundColor Green
Write-Host "Response: $($Response.Content)" -ForegroundColor Gray
# Cleanup
Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue
Remove-Item $TempZipBase -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Sync complete!" -ForegroundColor Green
exit 0
