# =============================================================================
# Start GenerateDeliveryReports web app (Kestrel, HTTP, port 5158)
# Run this script to start the app. Keep the window open to keep it running.
# Access from this machine : http://localhost:5158
# Access from other machines: http://<this-machine-ip>:5158
# =============================================================================

$ExePath = Join-Path $PSScriptRoot "GenerateDeliveryReports.exe"

if (-not (Test-Path $ExePath)) {
    Write-Host "ERROR: GenerateDeliveryReports.exe not found at $ExePath" -ForegroundColor Red
    Write-Host "Run deploy.ps1 first to publish the application." -ForegroundColor Yellow
    exit 1
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " GenerateDeliveryReports - Web App" -ForegroundColor Cyan
Write-Host " http://localhost:5158" -ForegroundColor Green
Write-Host " Press Ctrl+C to stop" -ForegroundColor Gray
Write-Host "========================================" -ForegroundColor Cyan

& $ExePath
