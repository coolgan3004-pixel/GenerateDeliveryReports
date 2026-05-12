# =============================================================================
# Deploy GenerateDeliveryReports (Kestrel) to a Windows machine via Task Scheduler
# =============================================================================
# Usage:
#   .\deploy.ps1 -TargetPath "\\MACHINE\C$\Apps\GenerateDeliveryReports"
#   .\deploy.ps1 -TargetPath "D:\Apps\GenerateDeliveryReports"  (local deploy)
#   .\deploy.ps1 -TargetPath "D:\Apps\GenerateDeliveryReports" -TaskName "MyAppTask"
#
# The script will:
#   1. Publish the app self-contained
#   2. Stop the existing Task Scheduler task and kill the process (if present)
#   3. Copy published files to TargetPath
#   4. Register (or update) a Task Scheduler task that runs the app at system startup
#   5. Start the task immediately
#
# The app binds to http://*:5158 (configured in Program.cs).
# The task runs as SYSTEM. Ensure OneDriveLocation paths are accessible to SYSTEM,
# or change -TaskName and configure a specific user account manually after deploy.
# =============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,

    [string]$Configuration = "Release",
    [string]$Runtime       = "win-x64",
    [string]$TaskName      = "GenerateDeliveryReports"
)

$ErrorActionPreference = "Stop"
$ProjectPath = Join-Path $PSScriptRoot "GenerateDeliveryReports\GenerateDeliveryReports.csproj"
$PublishDir = Join-Path $PSScriptRoot "publish"
$TemplateSrc = Join-Path $PSScriptRoot "GenerateDeliveryReports.Data\Templates"

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

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " GenerateDeliveryReports - Deploy Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Step 1: Clean previous publish output
if (Test-Path $PublishDir) {
    Write-Host "`n[1/7] Cleaning previous publish output..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $PublishDir
}
else {
    Write-Host "`n[1/7] No previous publish output to clean." -ForegroundColor Gray
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
Write-Host "`n[3/7] Copying report template..." -ForegroundColor Yellow
$TemplateDestDir = Join-Path $PublishDir "Templates"
if (-not (Test-Path $TemplateDestDir)) {
    New-Item -ItemType Directory -Path $TemplateDestDir | Out-Null
}
Copy-Item -Path (Join-Path $TemplateSrc "*") -Destination $TemplateDestDir -Force
Write-Host "Template copied to $TemplateDestDir" -ForegroundColor Green

# Step 4: Update template path in appsettings.json to be relative
Write-Host "`n[4/7] Updating appsettings.json for deployment..." -ForegroundColor Yellow
$AppSettingsPath = Join-Path $PublishDir "appsettings.json"

# Use targeted string replacement instead of ConvertFrom-Json → ConvertTo-Json to avoid
# PowerShell 5.1 silently dropping deeply-nested structures (e.g. the CSAT Clients array).
$content = Get-Content $AppSettingsPath -Raw
$newTemplatePath = 'Templates\\GlobalPayments-DeliveryQualitySummaryReport_Template.pptx'
$content = $content -replace '(?<="SprintMetricsReportTemplatePath"\s*:\s*")[^"]*(?=")', $newTemplatePath
# Clear the hardcoded dev-machine path so the app uses its built-in fallback (wwwroot/worker-summary.html)
$content = $content -replace '(?<="WorkerSummaryFilePath"\s*:\s*")[^"]*(?=")', ''
# Blank the dev-machine CommonFolderPath — must be set on the target machine
$content = $content -replace '(?<="CommonFolderPath"\s*:\s*")[^"]*(?=")', ''
[System.IO.File]::WriteAllText($AppSettingsPath, $content, [System.Text.Encoding]::UTF8)

# Verify CSAT section survived
if ($content -notmatch '"CSAT"') {
    Write-Host "ERROR: CSAT section is missing from appsettings.json after update. Aborting deploy." -ForegroundColor Red
    exit 1
}
if ($content -notmatch '"Clients"') {
    Write-Host "ERROR: CSAT Clients array is missing from appsettings.json after update. Aborting deploy." -ForegroundColor Red
    exit 1
}
Write-Host "appsettings.json updated." -ForegroundColor Green
Write-Host "  NOTE: Update 'OneDriveLocation' and 'CommonFolderPath' in appsettings.json on the target machine." -ForegroundColor Magenta

# Resolve exe path once — used in both Step 5 (process kill) and Step 7 (task registration)
$ExePath     = Join-Path $TargetPath "GenerateDeliveryReports.exe"
$ProcessName = [System.IO.Path]::GetFileNameWithoutExtension($ExePath)

# Step 5: Stop Task Scheduler task + kill the process to release file locks
Write-Host "`n[5/7] Stopping task '$TaskName' (if running)..." -ForegroundColor Yellow
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    # Kill the host process directly so all file handles are released before the copy
    Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Write-Host "Task stopped." -ForegroundColor Green
} else {
    Write-Host "Task not found -- will register after copy." -ForegroundColor Gray
}

# Step 6: Copy to target
Write-Host "`n[6/7] Deploying to $TargetPath ..." -ForegroundColor Yellow
if (-not (Test-Path $TargetPath)) {
    New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null
}
Copy-Item -Path (Join-Path $PublishDir "*") -Destination $TargetPath -Recurse -Force
Write-Host "Deployed successfully to $TargetPath" -ForegroundColor Green

# Step 7: Register (or update) the Task Scheduler task, then start it
Write-Host "`n[7/7] Configuring Task Scheduler task '$TaskName'..." -ForegroundColor Yellow
$action    = New-ScheduledTaskAction -Execute $ExePath -WorkingDirectory $TargetPath
$trigger   = New-ScheduledTaskTrigger -AtStartup
$settings  = New-ScheduledTaskSettingsSet `
                 -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
                 -RestartCount 3 `
                 -RestartInterval (New-TimeSpan -Minutes 1) `
                 -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal `
                 -UserId "SYSTEM" `
                 -LogonType ServiceAccount `
                 -RunLevel Highest

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    Set-ScheduledTask -TaskName $TaskName `
        -Action $action -Trigger $trigger -Settings $settings -Principal $principal | Out-Null
    Write-Host "Task updated." -ForegroundColor Green
} else {
    Register-ScheduledTask -TaskName $TaskName `
        -Description "GenerateDeliveryReports Kestrel web application" `
        -Action $action -Trigger $trigger -Settings $settings -Principal $principal | Out-Null
    Write-Host "Task registered." -ForegroundColor Green
}

Start-ScheduledTask -TaskName $TaskName
Write-Host "Task started. App should be reachable at http://<machine>:5158" -ForegroundColor Green

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " Deployment Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Post-deployment checklist on the target machine:" -ForegroundColor Yellow
Write-Host "  1. Edit $TargetPath\appsettings.json:" -ForegroundColor White
Write-Host "     - Set 'OneDriveLocation' to the local OneDrive sync path" -ForegroundColor White
Write-Host "     - Set 'CommonFolderPath' to a writable folder (e.g. D:\AppData\GenerateDeliveryReports)" -ForegroundColor White
Write-Host "       This folder will hold LogFiles\ and downloads\ sub-folders" -ForegroundColor White
Write-Host "  2. Verify the task is running:" -ForegroundColor White
Write-Host "     Get-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Gray
Write-Host "  3. Check the app is reachable:" -ForegroundColor White
Write-Host "     Invoke-WebRequest http://localhost:5158 -UseBasicParsing" -ForegroundColor Gray
Write-Host "  4. View application logs:" -ForegroundColor White
Write-Host ('     Get-Content "' + $TargetPath + '\LogFiles\log*.txt" -Tail 50') -ForegroundColor Gray
Write-Host "  5. To stop the app manually:" -ForegroundColor White
Write-Host "     Stop-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Gray
Write-Host ""
