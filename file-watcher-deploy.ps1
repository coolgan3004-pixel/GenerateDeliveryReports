# File Watcher - Monitor config paths and trigger deployment on changes
# Reads monitoring paths from appsettings.json
# Triggers deploy-appandfiles-azureappservice-scheduler.ps1 when changes detected

param(
    [Parameter(Mandatory=$true)][string]$AppSettingsPath,
    [Parameter(Mandatory=$true)][string]$AppId,
    [Parameter(Mandatory=$true)][string]$Password,
    [Parameter(Mandatory=$true)][string]$Tenant,
    [string]$BatchWaitSeconds = 30,
    [string]$DeploymentFolder = "",
    [string]$ResourceGroupName = "MakeThingsEasier-Deploy",
    [string]$AppServiceName = "DailyStatus"
)

$ErrorActionPreference = "Stop"

# Setup logging
$logDir = "C:\Logs"
$logPath = Join-Path $logDir ("filewatcher-" + (Get-Date -Format 'yyyy-MM-dd') + ".log")
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Log-Message {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Message"
    Write-Host $logEntry
    Add-Content -Path $logPath -Value $logEntry
}

Log-Message "========== File Watcher Started =========="

# Read appsettings
if (-not (Test-Path $AppSettingsPath)) {
    Log-Message "ERROR: appsettings.json not found at: $AppSettingsPath"
    exit 1
}

$appConfig = Get-Content $AppSettingsPath -Raw | ConvertFrom-Json
$oneDrivePath = $appConfig.AppSettings.OneDriveLocation
$metricsFolder = $appConfig.AppSettings.MetricsFolder.TrimStart('\').TrimStart('/')
$projectsFolder = "Projects"
$commonFolder = $appConfig.AppSettings.CommonFolderPath

# Construct monitoring paths
$pathCommonFolder = $commonFolder
$pathMetricsFolder = Join-Path $oneDrivePath $metricsFolder
$pathProjectsFolder = Join-Path $oneDrivePath $projectsFolder

Log-Message "Monitoring paths:"
Log-Message "  1. CommonFolder: $pathCommonFolder"
Log-Message "  2. MetricsFolder: $pathMetricsFolder"
Log-Message "  3. ProjectsFolder: $pathProjectsFolder"

# Verify paths exist
$paths = @($pathCommonFolder, $pathMetricsFolder, $pathProjectsFolder)
foreach ($path in $paths) {
    if (-not (Test-Path $path)) {
        Log-Message "WARNING: Path not found: $path"
    }
}

# Track last deployment time to avoid duplicates
$lastDeploymentTime = Get-Date
$deploymentScript = "C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports\deploy-appandfiles-azureappservice-scheduler.ps1"
$changeDetected = $false
$changeTime = $null

# Function to trigger deployment
function Trigger-Deployment {
    Log-Message "Change detected! Waiting $BatchWaitSeconds seconds for batch..."
    Start-Sleep -Seconds $BatchWaitSeconds

    Log-Message "Starting deployment..."
    try {
        & $deploymentScript `
          -AppSettingsPath $AppSettingsPath `
          -AppId $AppId `
          -Password $Password `
          -Tenant $Tenant `
          -ResourceGroupName $ResourceGroupName `
          -AppServiceName $AppServiceName `
          -DeploymentFolder $DeploymentFolder

        Log-Message "Deployment completed successfully"
        $global:lastDeploymentTime = Get-Date
    }
    catch {
        Log-Message "ERROR during deployment: $_"
    }
}

# Create file watchers
$watchers = @()
foreach ($path in $paths) {
    if (Test-Path $path) {
        $watcher = New-Object System.IO.FileSystemWatcher
        $watcher.Path = $path
        $watcher.IncludeSubdirectories = $true
        $watcher.EnableRaisingEvents = $false

        # Events for all file changes
        $changeScript = {
            if (-not $global:changeDetected) {
                $global:changeDetected = $true
                $global:changeTime = Get-Date
                Log-Message "File change detected in: $($Event.SourceEventArgs.FullPath)"
            }
        }

        Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $changeScript | Out-Null
        Register-ObjectEvent -InputObject $watcher -EventName "Changed" -Action $changeScript | Out-Null
        Register-ObjectEvent -InputObject $watcher -EventName "Deleted" -Action $changeScript | Out-Null

        $watchers += $watcher
        Log-Message "Watcher started for: $path"
    }
}

# Main loop
Log-Message "Watching for changes... Press Ctrl+C to stop"
$global:changeDetected = $false

while ($true) {
    Start-Sleep -Seconds 5

    if ($global:changeDetected) {
        $timeSinceChange = (Get-Date) - $global:changeTime

        if ($timeSinceChange.TotalSeconds -ge $BatchWaitSeconds) {
            Trigger-Deployment
            $global:changeDetected = $false
        }
    }
}

# Cleanup
foreach ($watcher in $watchers) {
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
}

Log-Message "File Watcher stopped"
exit 0
