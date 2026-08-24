# Monitor file changes and trigger Azure DevOps pipeline deployment
param(
    [Parameter(Mandatory=$true)][string]$AppSettingsPath,
    [Parameter(Mandatory=$true)][string]$PatToken,
    [int]$BatchWaitSeconds = 30
)

$ErrorActionPreference = "Continue"

$logDir = "C:\Logs"
$logPath = Join-Path $logDir ("filewatcher-trigger-" + (Get-Date -Format 'yyyy-MM-dd_HHmmss') + ".log")
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

Start-Transcript -Path $logPath -Append | Out-Null

Write-Host "========================================"
Write-Host "File Watcher - Pipeline Trigger Mode"
Write-Host "========================================"
Write-Host "Log file: $logPath"

if (-not (Test-Path $AppSettingsPath)) {
    Write-Host "ERROR: appsettings.json not found at: $AppSettingsPath"
    exit 1
}

Write-Host "Configuring Azure DevOps..."
$env:AZURE_DEVOPS_EXT_PAT = $PatToken
az devops configure --defaults organization=https://dev.azure.com/coolgan3004 project=PersonalRepo
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to configure Azure DevOps"
    exit 1
}
Write-Host "OK: Azure DevOps configured"
Write-Host ""

$appConfig = Get-Content $AppSettingsPath -Raw | ConvertFrom-Json
$oneDrivePath = $appConfig.AppSettings.OneDriveLocation
$metricsFolder = $appConfig.AppSettings.MetricsFolder.TrimStart('\').TrimStart('/')
$commonFolder = $appConfig.AppSettings.CommonFolderPath

$pathsToWatch = @(
    $commonFolder,
    (Join-Path $oneDrivePath $metricsFolder),
    (Join-Path $oneDrivePath "Projects")
)

Write-Host "Monitoring paths:"
foreach ($path in $pathsToWatch) {
    if (Test-Path $path) {
        Write-Host "  OK: $path"
    } else {
        Write-Host "  WARNING: $path (not found)"
    }
}

Write-Host ""
Write-Host "Watching for changes... Press Ctrl+C to stop"
Write-Host "Batch wait time: $BatchWaitSeconds seconds"
Write-Host ""

$watchers = @()
foreach ($path in $pathsToWatch) {
    if (Test-Path $path) {
        $watcher = New-Object System.IO.FileSystemWatcher
        $watcher.Path = $path
        $watcher.IncludeSubdirectories = $true
        $watcher.EnableRaisingEvents = $true

        Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action { $global:changeDetected = $true; $global:changeTime = Get-Date } | Out-Null
        Register-ObjectEvent -InputObject $watcher -EventName "Changed" -Action { $global:changeDetected = $true; $global:changeTime = Get-Date } | Out-Null
        Register-ObjectEvent -InputObject $watcher -EventName "Deleted" -Action { $global:changeDetected = $true; $global:changeTime = Get-Date } | Out-Null

        $watchers += $watcher
    }
}

$global:changeDetected = $false

while ($true) {
    Start-Sleep -Seconds 5

    if ($global:changeDetected) {
        $timeSinceChange = (Get-Date) - $global:changeTime

        if ($timeSinceChange.TotalSeconds -ge $BatchWaitSeconds) {
            Write-Host ""
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Changes detected! Triggering pipeline..."
            Write-Host ""

            az pipelines run --id 5 --project "PersonalRepo" --branch main --org https://dev.azure.com/coolgan3004

            if ($LASTEXITCODE -eq 0) {
                Write-Host "OK: Pipeline triggered"
            } else {
                Write-Host "ERROR: Failed to trigger pipeline (exit code: $LASTEXITCODE)"
            }

            Write-Host ""
            Write-Host "Watching for changes... Press Ctrl+C to stop"
            $global:changeDetected = $false
        }
    }
}

foreach ($watcher in $watchers) {
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
}

Stop-Transcript | Out-Null
exit 0
