# =============================================================================
# Deploy GenerateDeliveryReports to a remote Windows machine
# =============================================================================
# Usage:
#   .\deploy.ps1 -TargetPath "\\MACHINE\C$\Apps\GenerateDeliveryReports"
#   .\deploy.ps1 -TargetPath "D:\Apps\GenerateDeliveryReports"  (local deploy)
# =============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetPath,

    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$IISSiteName = "GenerateDeliveryReports"  # Set to empty string to skip IIS restart
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
$settings = Get-Content $AppSettingsPath -Raw | ConvertFrom-Json

# Set template path relative to the exe
$settings.AppSettings.SprintMetricsReportTemplatePath = "Templates\GlobalPayments-DeliveryQualitySummaryReport_Template.pptx"

$settings | ConvertTo-Json -Depth 10 | Set-Content $AppSettingsPath -Encoding UTF8
Write-Host "appsettings.json updated." -ForegroundColor Green
Write-Host "  NOTE: Update 'OneDriveLocation' in appsettings.json on the target machine." -ForegroundColor Magenta

# Step 5: Stop IIS site
if ($IISSiteName) {
    Write-Host "`n[5/7] Stopping IIS site '$IISSiteName'..." -ForegroundColor Yellow
    & iisreset /stop
    Write-Host "IIS stopped." -ForegroundColor Green
} else {
    Write-Host "`n[5/7] No IIS site specified -- skipping IIS stop." -ForegroundColor Gray
}

# Step 6: Copy to target
Write-Host "`n[6/7] Deploying to $TargetPath ..." -ForegroundColor Yellow
if (-not (Test-Path $TargetPath)) {
    New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null
}
Copy-Item -Path (Join-Path $PublishDir "*") -Destination $TargetPath -Recurse -Force
Write-Host "Deployed successfully to $TargetPath" -ForegroundColor Green

# Step 7: Start IIS site
if ($IISSiteName) {
    Write-Host "`n[7/7] Starting IIS site '$IISSiteName'..." -ForegroundColor Yellow
    & iisreset /start
    Write-Host "IIS started." -ForegroundColor Green
} else {
    Write-Host "`n[7/7] No IIS site specified -- skipping IIS start." -ForegroundColor Gray
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " Deployment Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps on the target machine:" -ForegroundColor Yellow
Write-Host "  1. Edit appsettings.json:" -ForegroundColor White
Write-Host "     - Set 'OneDriveLocation' to the local OneDrive sync path" -ForegroundColor White
Write-Host "  2. Run the app:" -ForegroundColor White
Write-Host "     - Direct:  .\GenerateDeliveryReports.exe" -ForegroundColor White
Write-Host "     - Service: sc.exe create DeliveryReports binPath=`"$TargetPath\GenerateDeliveryReports.exe`" start=auto" -ForegroundColor White
Write-Host ""
