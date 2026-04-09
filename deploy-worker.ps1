# =============================================================================
# Deploy GenerateDeliveryReports.Worker as a Windows Service
# =============================================================================
# Usage:
#   .\deploy-worker.ps1 -TargetPath "\\MACHINE\C$\Services\DeliveryReportsWorker"
#   .\deploy-worker.ps1 -TargetPath "D:\Services\DeliveryReportsWorker"  (local)
#
# Pipeline usage (set variables in Azure DevOps):
#   .\deploy-worker.ps1 -TargetPath "$(TargetPath)" -ServiceName "$(ServiceName)"
# =============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,

    [string]$ServiceName    = "DeliveryReportsWorker",
    [string]$ServiceDisplay = "Delivery Reports Worker",
    [string]$Configuration  = "Release",
    [string]$Runtime        = "win-x64"
)

$ErrorActionPreference = "Stop"
$ProjectPath  = Join-Path $PSScriptRoot "GenerateDeliveryReports.Worker\GenerateDeliveryReports.Worker.csproj"
$PublishDir   = Join-Path $PSScriptRoot "publish-worker"
$TemplateSrc  = Join-Path $PSScriptRoot "GenerateDeliveryReports.Data\Templates"

# Locate dotnet.exe
$dotnetCmd = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnetCmd) {
    foreach ($p in @(
        "$env:ProgramFiles\dotnet\dotnet.exe",
        "${env:ProgramFiles(x86)}\dotnet\dotnet.exe",
        "$env:LOCALAPPDATA\Microsoft\dotnet\dotnet.exe"
    )) {
        if (Test-Path $p) { $dotnetCmd = $p; break }
    }
    if (-not $dotnetCmd) {
        Write-Host "ERROR: 'dotnet' not found. Install .NET SDK or add it to PATH." -ForegroundColor Red
        exit 1
    }
}
$dotnet = if ($dotnetCmd -is [string]) { $dotnetCmd } else { $dotnetCmd.Source }
Write-Host "Using dotnet: $dotnet" -ForegroundColor Gray

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " GenerateDeliveryReports.Worker - Deploy   " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# Step 1: Clean previous publish output
if (Test-Path $PublishDir) {
    Write-Host "`n[1/6] Cleaning previous publish output..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $PublishDir
}

# Step 2: Publish self-contained
Write-Host "`n[2/6] Publishing ($Configuration | $Runtime | self-contained)..." -ForegroundColor Yellow
& $dotnet publish $ProjectPath -c $Configuration -r $Runtime --self-contained -o $PublishDir
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Publish failed." -ForegroundColor Red
    exit 1
}
Write-Host "Publish succeeded." -ForegroundColor Green

# Step 3: Copy PPTX template
Write-Host "`n[3/6] Copying report template..." -ForegroundColor Yellow
$TemplateDest = Join-Path $PublishDir "Templates"
if (-not (Test-Path $TemplateDest)) { New-Item -ItemType Directory -Path $TemplateDest | Out-Null }
Copy-Item -Path (Join-Path $TemplateSrc "*") -Destination $TemplateDest -Force
Write-Host "Templates copied." -ForegroundColor Green

# Step 4: Stop service if running
Write-Host "`n[4/6] Stopping service '$ServiceName' (if running)..." -ForegroundColor Yellow
$svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($svc) {
    if ($svc.Status -ne 'Stopped') {
        Stop-Service -Name $ServiceName -Force
        $svc.WaitForStatus('Stopped', (New-TimeSpan -Seconds 30))
        Write-Host "Service stopped." -ForegroundColor Green
    } else {
        Write-Host "Service already stopped." -ForegroundColor Gray
    }
} else {
    Write-Host "Service not yet registered — will create after copy." -ForegroundColor Gray
}

# Step 5: Copy to target
Write-Host "`n[5/6] Deploying to $TargetPath ..." -ForegroundColor Yellow
if (-not (Test-Path $TargetPath)) { New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null }
Copy-Item -Path (Join-Path $PublishDir "*") -Destination $TargetPath -Recurse -Force
Write-Host "Files deployed." -ForegroundColor Green

# Step 6: Register or start service
$ExePath = Join-Path $TargetPath "GenerateDeliveryReports.Worker.exe"
Write-Host "`n[6/6] Registering / starting service '$ServiceName'..." -ForegroundColor Yellow

$svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $svc) {
    sc.exe create $ServiceName binPath= "`"$ExePath`"" DisplayName= "$ServiceDisplay" start= auto | Out-Null
    sc.exe description $ServiceName "Generates delivery quality summary reports on a schedule." | Out-Null
    Write-Host "Service registered." -ForegroundColor Green
} else {
    sc.exe config $ServiceName binPath= "`"$ExePath`"" | Out-Null
    Write-Host "Service binary path updated." -ForegroundColor Green
}

Start-Service -Name $ServiceName
Write-Host "Service started." -ForegroundColor Green

# Summary
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " Deployment Complete!" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Post-deployment checklist on target machine:" -ForegroundColor Yellow
Write-Host "  1. Edit $TargetPath\appsettings.json:" -ForegroundColor White
Write-Host "     - Set 'OneDriveLocation' to the local OneDrive sync path" -ForegroundColor White
Write-Host "     - Set 'WorkerSummaryFilePath' to where the Blazor wwwroot is" -ForegroundColor White
Write-Host "     - Set 'EmailSettings.Password' (or set SENDGRID_API_KEY env var)" -ForegroundColor White
Write-Host "  2. Check service status:" -ForegroundColor White
Write-Host "     Get-Service $ServiceName" -ForegroundColor Gray
Write-Host "  3. View logs:" -ForegroundColor White
Write-Host "     Get-Content $TargetPath\LogFiles\workerlog*.txt -Tail 50" -ForegroundColor Gray
Write-Host ""
