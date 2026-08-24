# Deploy GenerateDeliveryReports App + Files to Azure App Service
# For scheduled deployments every 30 minutes

param(
    [Parameter(Mandatory=$true)][string]$AppSettingsPath,
    [Parameter(Mandatory=$true)][string]$AppId,
    [Parameter(Mandatory=$true)][string]$Password,
    [Parameter(Mandatory=$true)][string]$Tenant,
    [string]$ResourceGroupName = "MakeThingsEasier-Deploy",
    [string]$AppServiceName = "DailyStatus",
    [string]$DeploymentFolder = "",
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x86",
    [switch]$Watch = $false,
    [int]$BatchWaitSeconds = 30
)

# Don't stop on warnings, only on errors
$ErrorActionPreference = "Continue"

# Setup logging
$logDir = "C:\Logs"
$logPath = Join-Path $logDir ("deploy-appandfiles-" + (Get-Date -Format 'yyyy-MM-dd_HHmmss') + ".log")
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# Start logging
Start-Transcript -Path $logPath -Append | Out-Null

Write-Host "========================================"
Write-Host "GenerateDeliveryReports App Deployment"
Write-Host "========================================"
Write-Host "Log file: $logPath"

if (-not (Test-Path $AppSettingsPath)) {
    Write-Host "ERROR: appsettings.json not found at: $AppSettingsPath"
    exit 1
}

if ($Watch) {
    Write-Host "WATCH MODE ENABLED - Will redeploy on file changes"
    Write-Host "Batch wait time: $BatchWaitSeconds seconds"
}

Write-Host "[1/6] Authenticating with Azure..."
$result = az login --service-principal -u $AppId -p $Password --tenant $Tenant
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Authentication failed"
    exit 1
}
Write-Host "OK: Authenticated"

Write-Host "[2/6] Preparing CommonFiles..."
$GenerateScript = "C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports\generate-sprint-reports-availability.ps1"
$appConfig = Get-Content $AppSettingsPath -Raw | ConvertFrom-Json
$oneDriveBase = $appConfig.AppSettings.OneDriveLocation.TrimEnd('\')

$TempBase = Join-Path $env:TEMP ("ZipTemp_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
$TempCommonFiles = Join-Path $TempBase "CommonFiles"
$FilesDestDir = Join-Path $TempCommonFiles "Files"
$TemplatesDestDir = Join-Path $TempCommonFiles "Templates"

New-Item -ItemType Directory -Path $FilesDestDir -Force | Out-Null
New-Item -ItemType Directory -Path $TemplatesDestDir -Force | Out-Null

$metricsFolder = $appConfig.AppSettings.MetricsFolder.TrimStart('\').TrimStart('/')
$dataFolder = $appConfig.AppSettings.ReportAndDataFolder.TrimStart('\').TrimStart('/')
$xlsxCount = 0

foreach ($project in $appConfig.AppSettings.Projects) {
    foreach ($metricsPath in $project.MetricsSheetPath) {
        if (-not $metricsPath) {
            continue
        }
        $srcFull = Join-Path $oneDriveBase "$metricsFolder\$metricsPath"
        $dstFull = Join-Path $FilesDestDir "$metricsFolder\$metricsPath"
        if (Test-Path $srcFull) {
            $dstDir = Split-Path -Parent $dstFull
            New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
            Copy-Item -Path $srcFull -Destination $dstFull -Force
            $xlsxCount = $xlsxCount + 1
        }
    }

    if ($project.DataFileName) {
        $srcFull = Join-Path $oneDriveBase "$dataFolder\$($project.DataFileName)"
        $dstFull = Join-Path $FilesDestDir "$dataFolder\$($project.DataFileName)"
        if (Test-Path $srcFull) {
            $dstDir = Split-Path -Parent $dstFull
            New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
            Copy-Item -Path $srcFull -Destination $dstFull -Force
            $xlsxCount = $xlsxCount + 1
        }
    }
}

Write-Host "OK: Bundled $xlsxCount xlsx files"

$templatePath = $appConfig.AppSettings.SprintMetricsReportTemplatePath
if ($templatePath -and (Test-Path $templatePath)) {
    $templateDir = Split-Path -Parent $templatePath
    if (Test-Path $templateDir) {
        Copy-Item -Path (Join-Path $templateDir "*") -Destination $TemplatesDestDir -Force -Recurse
        Write-Host "OK: Template bundled"
    }
}

$csatFolder = $appConfig.AppSettings.CSAT.CSATFolder.TrimStart('\').TrimStart('/')
$csatSeen = @{}
foreach ($client in $appConfig.AppSettings.CSAT.Clients) {
    $fileName = $client.ClientSurveyFilePath.Trim()
    if (-not $fileName -or $csatSeen.ContainsKey($fileName)) {
        continue
    }
    $csatSeen[$fileName] = $true
    $srcFull = Join-Path $oneDriveBase "$csatFolder\$fileName"
    $dstFull = Join-Path $FilesDestDir "$csatFolder\$fileName"
    if (Test-Path $srcFull) {
        $dstDir = Split-Path -Parent $dstFull
        New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
        Copy-Item -Path $srcFull -Destination $dstFull -Force
    }
}

Write-Host "Generating JSON..."
$jsonOutputPath = Join-Path $FilesDestDir "SprintReportsAvailability.json"
if (Test-Path $GenerateScript) {
    & $GenerateScript -AppSettingsPath $AppSettingsPath -OutputJson $jsonOutputPath
    Write-Host "OK: JSON generated"
}

$WorkerHistoryDir = Join-Path $TempCommonFiles "WorkerHistory"
if ($DeploymentFolder -and (Test-Path $DeploymentFolder)) {
    New-Item -ItemType Directory -Path $WorkerHistoryDir -Force | Out-Null
    $htmlFile = Get-ChildItem $DeploymentFolder -Filter "run-workerhistory.html" -Recurse | Select-Object -First 1
    if ($htmlFile) {
        Copy-Item -Path $htmlFile.FullName -Destination $WorkerHistoryDir -Force
        Write-Host "OK: Worker HTML bundled"
    }
    $jsonFile = Get-ChildItem $DeploymentFolder -Filter "run-workerhistory.json" -Recurse | Select-Object -First 1
    if ($jsonFile) {
        Copy-Item -Path $jsonFile.FullName -Destination $WorkerHistoryDir -Force
        Write-Host "OK: Worker JSON bundled"
    }
}

Write-Host "[3/6] Publishing app..."
$projectPath = "C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports\GenerateDeliveryReports\GenerateDeliveryReports.csproj"
$publishPath = Join-Path $TempBase "publish"

dotnet publish $projectPath -c $Configuration -r $Runtime -o $publishPath --self-contained
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Publish failed"
    exit 1
}
Write-Host "OK: App published"

Write-Host "[4/6] Packaging CommonFiles..."
$publishCommonFilesPath = Join-Path $publishPath "CommonFiles"
if (Test-Path $publishCommonFilesPath) {
    Remove-Item $publishCommonFilesPath -Recurse -Force
}
Copy-Item -Path $TempCommonFiles -Destination $publishCommonFilesPath -Recurse -Force
Write-Host "OK: CommonFiles packaged"

Write-Host "[5/6] Creating deployment package..."
$zipPath = Join-Path $env:TEMP ("DailyStatus_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".zip")
Compress-Archive -Path $publishPath -DestinationPath $zipPath -Force
Write-Host "OK: Zip created"

Write-Host "[6/6] Deploying to Azure..."
Write-Host "Resource Group: $ResourceGroupName"
Write-Host "App Service: $AppServiceName"
Write-Host "Zip Path: $zipPath"
Write-Host "Zip exists: $(Test-Path $zipPath)"

$deployOutput = az webapp deployment source config-zip --resource-group $ResourceGroupName --name $AppServiceName --src $zipPath 2>&1 | Where-Object { $_ -notmatch "UserWarning" }

if ($LASTEXITCODE -eq 0) {
    Write-Host "OK: Deployment successful"
    Write-Host "Deployment output: $deployOutput"
} else {
    Write-Host "ERROR: Deployment failed with exit code: $LASTEXITCODE"
    Write-Host "Output: $deployOutput"
    exit 1
}

Write-Host "Cleaning up..."
Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
Remove-Item $TempBase -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "========================================"
Write-Host "Deployment Complete!"
Write-Host "========================================"
Write-Host "Log saved to: $logPath"

# Start file watcher if requested
if ($Watch) {
    Write-Host ""
    Write-Host "========================================"
    Write-Host "Starting File Watcher Mode"
    Write-Host "========================================"

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
            Write-Host "  - $path"
        } else {
            Write-Host "  - $path (WARNING: not found)"
        }
    }

    $changeDetected = $false
    $changeTime = $null
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

    Write-Host "Watching for changes... Press Ctrl+C to stop"
    $global:changeDetected = $false

    while ($true) {
        Start-Sleep -Seconds 5

        if ($global:changeDetected) {
            $timeSinceChange = (Get-Date) - $global:changeTime

            if ($timeSinceChange.TotalSeconds -ge $BatchWaitSeconds) {
                Write-Host ""
                Write-Host "File changes detected! Redeploying..."
                Write-Host ""

                & $PSCommandPath `
                  -AppSettingsPath $AppSettingsPath `
                  -AppId $AppId `
                  -Password $Password `
                  -Tenant $Tenant `
                  -ResourceGroupName $ResourceGroupName `
                  -AppServiceName $AppServiceName `
                  -DeploymentFolder $DeploymentFolder `
                  -Watch:$false

                $global:changeDetected = $false
            }
        }
    }

    foreach ($watcher in $watchers) {
        $watcher.EnableRaisingEvents = $false
        $watcher.Dispose()
    }
}

Stop-Transcript | Out-Null
exit 0
