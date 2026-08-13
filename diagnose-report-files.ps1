# =============================================================================
# diagnose-report-files.ps1
# =============================================================================
# Diagnostic script to find all PPTX files in project folders and identify
# mismatches with the expected filename pattern.
#
# Usage:
#   .\diagnose-report-files.ps1
# =============================================================================

param(
    [string]$AppSettingsPath = $null
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# 1. Resolve paths
# ---------------------------------------------------------------------------
$scriptDir = $PSScriptRoot

if (-not $AppSettingsPath) {
    $AppSettingsPath = Join-Path $scriptDir "GenerateDeliveryReports\appsettings.json"
}
if (-not (Test-Path $AppSettingsPath)) {
    Write-Host "ERROR: appsettings.json not found at $AppSettingsPath" -ForegroundColor Red
    exit 1
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Report Files Diagnostic" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# 2. Load appsettings.json
# ---------------------------------------------------------------------------
$settings    = Get-Content $AppSettingsPath -Raw | ConvertFrom-Json
$appSettings = $settings.AppSettings

$oneDriveBase = "$($appSettings.OneDriveLocation)".TrimEnd('\')
$dataFolder   = "$($appSettings.ReportAndDataFolder)".TrimStart('\').TrimStart('/')
$projects     = $appSettings.Projects

if (-not (Test-Path $oneDriveBase)) {
    Write-Host "ERROR: OneDriveLocation path not found: $oneDriveBase" -ForegroundColor Red
    exit 1
}

Write-Host "OneDrive Base: $oneDriveBase" -ForegroundColor Gray
Write-Host ""

# ---------------------------------------------------------------------------
# 3. Helper: Find files by fuzzy matching
# ---------------------------------------------------------------------------
function Find-SimilarFile([string]$projectName, [string]$sprintName, [string]$projectDir) {
    if (-not (Test-Path $projectDir)) {
        return $null
    }

    $pptxFiles = Get-ChildItem -Path $projectDir -Filter "*.pptx" -File -ErrorAction SilentlyContinue

    if ($pptxFiles.Count -eq 0) {
        return $null
    }

    # Exact match first
    foreach ($file in $pptxFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        if ($baseName -like "*$projectName*" -and $baseName -like "*$sprintName*") {
            return $file
        }
    }

    # Partial match: project name + something with sprint
    foreach ($file in $pptxFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        # Check if it contains the sprint name (more flexible)
        if ($baseName -like "*$sprintName*") {
            return $file
        }
    }

    # Very fuzzy: just check if project name is there
    foreach ($file in $pptxFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        if ($baseName -like "*$projectName*") {
            return $file
        }
    }

    return $null
}

# ---------------------------------------------------------------------------
# 4. Scan each project
# ---------------------------------------------------------------------------
$totalExpected = 0
$totalFound = 0
$totalMissing = 0
$mismatchExamples = @()

foreach ($project in $projects) {
    $projectName  = $project.ProjectName
    $dataFilePath = Join-Path $oneDriveBase (Join-Path $dataFolder $project.DataFileName)
    $projectDir   = Split-Path $dataFilePath -Parent

    Write-Host "[$projectName]" -ForegroundColor White
    Write-Host "  Folder: $projectDir" -ForegroundColor Gray

    if (-not (Test-Path $projectDir)) {
        Write-Host "  SKIP -- folder not found" -ForegroundColor Yellow
        continue
    }

    # List all PPTX files in the folder
    $pptxFiles = Get-ChildItem -Path $projectDir -Filter "*.pptx" -File -ErrorAction SilentlyContinue

    if ($pptxFiles.Count -eq 0) {
        Write-Host "  No PPTX files found in this folder" -ForegroundColor Yellow
        Write-Host ""
        continue
    }

    Write-Host "  Found $($pptxFiles.Count) PPTX file(s):" -ForegroundColor Green
    foreach ($file in $pptxFiles) {
        Write-Host "    - $($file.Name)" -ForegroundColor Cyan
    }
    Write-Host ""
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Analysis Complete" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Review the filenames above" -ForegroundColor Yellow
Write-Host "2. Note any patterns that don't match the expected format:" -ForegroundColor Yellow
Write-Host "   GlobalPayments-{ProjectName}-DeliveryQualitySummaryReport-{SprintName}.pptx" -ForegroundColor Yellow
Write-Host "3. Share examples of mismatches for fallback logic implementation" -ForegroundColor Yellow
Write-Host ""
