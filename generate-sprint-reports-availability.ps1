# =============================================================================
# generate-sprint-reports-availability.ps1
# =============================================================================
# Generates a matrix JSON file showing products and reports sent per day.
# Color-codes reports based on lag days from sprint end date to report creation.
#   - Green: 0 days lag (same day)
#   - Yellow: 1-2 days lag
#   - Red: 3+ days lag
#   - Missing: No report found
#
# Requirements:
#   ImportExcel module -- install once per machine:
#   Install-Module ImportExcel -Scope CurrentUser -Force
#
# Usage:
#   .\generate-sprint-reports-availability.ps1
#   .\generate-sprint-reports-availability.ps1 -AppSettingsPath "C:\path\to\appsettings.json"
#   .\generate-sprint-reports-availability.ps1 -OutputJson "C:\Deployments\SprintReportsAvailability.json"
#   .\generate-sprint-reports-availability.ps1 -FromDate "2026-06-01"
# =============================================================================

param(
    # Path to the source appsettings.json (dev machine copy, not the published one)
    [string]$AppSettingsPath = $null,

    # Where to write the output JSON file
    [string]$OutputJson = $null,

    # Ignore sprints whose end date is before this date (default: Apr 1 2026)
    [string]$FromDate = "2026-04-01"
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# 1. Prerequisites
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ERROR: ImportExcel module not found." -ForegroundColor Red
    Write-Host "Install it with:  Install-Module ImportExcel -Scope CurrentUser -Force" -ForegroundColor Yellow
    exit 1
}
Import-Module ImportExcel -ErrorAction Stop

# ---------------------------------------------------------------------------
# 2. Resolve paths
# ---------------------------------------------------------------------------
$scriptDir = $PSScriptRoot

if (-not $AppSettingsPath) {
    $AppSettingsPath = Join-Path $scriptDir "GenerateDeliveryReports\appsettings.json"
}
if (-not (Test-Path $AppSettingsPath)) {
    Write-Host "ERROR: appsettings.json not found at $AppSettingsPath" -ForegroundColor Red
    exit 1
}

if (-not $OutputJson) {
    $OutputJson = Join-Path $scriptDir "SprintReportsAvailability.json"
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Sprint Reports Availability -- Matrix Generator" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
$fromDateParsed = [DateTime]::Parse($FromDate)

Write-Host "  Settings : $AppSettingsPath" -ForegroundColor Gray
Write-Host "  Output   : $OutputJson"      -ForegroundColor Gray
Write-Host "  From     : $($fromDateParsed.ToString('dd-MMM-yyyy')) (sprints ending before this date are skipped)" -ForegroundColor Gray
Write-Host ""

# ---------------------------------------------------------------------------
# 3. Load appsettings.json
# ---------------------------------------------------------------------------
$settings    = Get-Content $AppSettingsPath -Raw | ConvertFrom-Json
$appSettings = $settings.AppSettings

$oneDriveBase = "$($appSettings.OneDriveLocation)".TrimEnd('\')
$dataFolder   = "$($appSettings.ReportAndDataFolder)".TrimStart('\').TrimStart('/')
$metricsBase  = "$($appSettings.MetricsFolder)".TrimStart('\').TrimStart('/')
$projects     = $appSettings.Projects

if (-not $oneDriveBase) {
    Write-Host "ERROR: OneDriveLocation is empty in appsettings.json" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $oneDriveBase)) {
    Write-Host "ERROR: OneDriveLocation path not found: $oneDriveBase" -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# 4. Helper functions
# ---------------------------------------------------------------------------

# Equivalent of GetRecentlyModifiedSimilarFile() from FileExtensions.cs:
# Finds FileName.xlsx or FileName (1).xlsx -- picks the most recently modified.
function Get-RecentSimilarFile([string]$filePath) {
    $dir      = Split-Path $filePath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
    $ext      = [System.IO.Path]::GetExtension($filePath)

    if (-not (Test-Path $dir -PathType Container)) { return $null }

    $escaped = [regex]::Escape($baseName)
    return Get-ChildItem -Path $dir -Filter "$baseName*$ext" -File -ErrorAction SilentlyContinue |
           Where-Object {
               $n = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
               # Exact match OR "Name (N)" duplicate pattern (OneDrive copy behaviour)
               ($n -eq $baseName) -or ($n -match "^$escaped\s*\(\d+\)$")
           } |
           Where-Object { -not $_.Name.StartsWith('~$') } |
           Sort-Object LastWriteTime -Descending |
           Select-Object -First 1
}

# Replicates the date-format list from DataProcessor.GetSprintNamesWithDate()
function Parse-SprintDate([string]$dateStr) {
    $formats = @(
        # ISO format (from Excel conversion)
        "yyyy-MM-dd",
        # With timestamps
        "MM/dd/yyyy HH:mm:ss","dd/MM/yyyy HH:mm:ss","yyyy-MM-dd HH:mm:ss",
        "d-MMM-yyyy HH:mm:ss","dd-MMM-yyyy HH:mm:ss",
        # Original formats
        "d-MMM-yyyy","dd-MMM-yyyy",
        "MMM-d-yyyy","MMM-dd-yyyy",
        "d-MMMM-yyyy","dd-MMMM-yyyy",
        "dd/MM/yyyy","d/M/yyyy","MM/dd/yyyy","M/d/yyyy",
        "dd-MM-yyyy","d-M-yyyy","dd.MM.yyyy",
        # 2-digit year formats
        "d-MMM-yy","dd-MMM-yy",
        "MMM-d-yy","MMM-dd-yy",
        "d-MMMM-yy","dd-MMMM-yy",
        "dd/MM/yy","d/M/yy","MM/dd/yy","M/d/yy",
        "dd-MM-yy","d-M-yy","dd.MM.yy"
    )
    $parsed = [DateTime]::MinValue
    foreach ($fmt in $formats) {
        if ([DateTime]::TryParseExact(
                $dateStr, $fmt,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$parsed)) {
            return $parsed
        }
    }
    return $null
}

# Determine status and color based on lag days
function Get-StatusAndColor([int]$lagDays) {
    if ($lagDays -eq 0) {
        return "Green"
    } elseif ($lagDays -le 2) {
        return "Yellow"
    } else {
        return "Red"
    }
}

# Convert Excel serial numbers to ISO date strings, or return formatted dates as-is
function Format-ExcelDate([string]$dateValue) {
    if (-not $dateValue -or $dateValue -eq 'nan') {
        return $null
    }

    # If it's already a formatted date string (contains letters), return as-is
    if ($dateValue -match '[A-Za-z]') {
        return $dateValue.Trim()
    }

    # If it's a number, assume it's an Excel serial number
    if ([int]::TryParse($dateValue, [ref]$null)) {
        try {
            # Excel serial date: Jan 1, 1900 = 1
            $excelDate = [int]$dateValue
            $baseDate = [DateTime]::new(1900, 1, 1)
            $actualDate = $baseDate.AddDays($excelDate - 1)
            return $actualDate.ToString("yyyy-MM-dd")
        }
        catch {
            return $null
        }
    }

    return $dateValue.Trim()
}

# ---------------------------------------------------------------------------
# 5. Process each project
# ---------------------------------------------------------------------------
$matrix = [System.Collections.Generic.List[object]]::new()

foreach ($project in $projects) {
    $projectName        = $project.ProjectName
    $metricsSheetPaths  = $project.MetricsSheetPath  # This is an array
    $dataFilePath       = Join-Path $oneDriveBase (Join-Path $dataFolder $project.DataFileName)
    $projectDir         = Split-Path $dataFilePath -Parent

    Write-Host "[$projectName]" -ForegroundColor White

    if (-not $metricsSheetPaths -or $metricsSheetPaths.Count -eq 0) {
        Write-Host "  SKIP -- no MetricsSheetPath configured." -ForegroundColor Yellow
        continue
    }

    # Process each metrics sheet file for this project
    $rows = @()
    foreach ($metricSheetPath in $metricsSheetPaths) {
        $fullMetricsPath = Join-Path $oneDriveBase (Join-Path $metricsBase $metricSheetPath)
        Write-Host "  MetricsSheet: $fullMetricsPath" -ForegroundColor Gray

        # Find the actual file (handles OneDrive duplicate "File (1).xlsx" copies)
        $actualFile = Get-RecentSimilarFile $fullMetricsPath
        if (-not $actualFile) {
            Write-Host "  SKIP -- metrics file not found: $metricSheetPath" -ForegroundColor Yellow
            continue
        }
        if ($actualFile.FullName -ne $fullMetricsPath) {
            Write-Host "  Using   : $($actualFile.FullName)" -ForegroundColor Gray
        }

        # Read the Dashboard worksheet - columns 2, 3, 4 (Sprint Name, Start Date, End Date)
        try {
            $sheetRows = Import-Excel -Path $actualFile.FullName -WorksheetName "Dashboard" `
                                     -ErrorAction Stop -NoHeader

            # Skip header row and extract columns 2, 3, 4
            for ($i = 1; $i -lt $sheetRows.Count; $i++) {
                $row = $sheetRows[$i]
                $sprintName = $row.P2   # Column 2 (Sprint Name)
                $startDate  = $row.P3   # Column 3 (Start Date)
                $endDate    = $row.P4   # Column 4 (End Date)

                if ($sprintName -and $sprintName -ne 'nan' -and $sprintName -ne 'Sprint #') {
                    # Convert Excel serial numbers to dates if needed
                    $startDateFormatted = Format-ExcelDate $startDate
                    $endDateFormatted   = Format-ExcelDate $endDate

                    $rows += [PSCustomObject]@{
                        SprintName = "$sprintName".Trim()
                        StartDate  = $startDateFormatted
                        EndDate    = $endDateFormatted
                    }
                }
            }
        }
        catch {
            Write-Host "  SKIP -- could not open Dashboard sheet: $_" -ForegroundColor Yellow
            continue
        }
    }

    if ($rows.Count -eq 0) {
        Write-Host "  No sprints found in metrics sheets." -ForegroundColor Gray
        continue
    }

    $entries      = [System.Collections.Generic.List[object]]::new()
    $sprintCount  = 0
    $parseErrorCount = 0
    $missingCount = 0

    foreach ($row in $rows) {
        $sprintName = $row.SprintName
        $startDateStr = $row.StartDate
        $endDateStr   = $row.EndDate

        if (-not $sprintName -or $sprintName -eq '' -or $sprintName -eq 'nan') { continue }

        # ---- Parse start / end dates from separate columns ----
        $startDate  = $null
        $endDate    = $null
        $parseError = $null

        if ($startDateStr -and $startDateStr -ne '' -and $startDateStr -ne 'nan') {
            $startDate = Parse-SprintDate $startDateStr
            if (-not $startDate) { $parseError = "Unrecognised start date format: '$startDateStr'" }
        } else {
            $parseError = "Start date is missing or empty"
        }

        if ($endDateStr -and $endDateStr -ne '' -and $endDateStr -ne 'nan') {
            $endDate = Parse-SprintDate $endDateStr
            if (-not $endDate) { $parseError = if ($parseError) { "$parseError; unrecognised end date format: '$endDateStr'" }
                                              else              { "Unrecognised end date format: '$endDateStr'" } }
        } else {
            $parseError = if ($parseError) { "$parseError; end date is missing or empty" }
                         else              { "End date is missing or empty" }
        }

        # ---- Parse error: log to console and record in JSON ----
        if ($parseError) {
            Write-Host ("  [??] {0}" -f $sprintName) -ForegroundColor Red
            Write-Host ("       PARSE ERROR: {0}" -f $parseError) -ForegroundColor Red

            $entries.Add([PSCustomObject]@{
                sprintName          = $sprintName.Trim()
                sprintStartDate     = $null
                sprintEndDate       = $null
                reportCreatedDate   = $null
                lagDays             = $null
                status              = "Missing"
                parseError          = $parseError
            })
            $parseErrorCount++
            continue
        }

        # ---- Intentional cutoff: silently skip sprints before FromDate ----
        if ($endDate.Date -lt $fromDateParsed.Date) { continue }

        # Use the sprint name directly from the dashboard sheet
        $sprintNameFormatted = $sprintName.Trim()

        # Only accept exact filename match - no fallback strategies
        # Teams must provide files with the correct standard naming convention
        $reportExists = $false
        $reportFile = $null
        $actualFileName = $null

        # Exact match - standard naming convention only
        $pptxName = "GlobalPayments-$projectName-DeliveryQualitySummaryReport-$sprintNameFormatted.pptx"
        $pptxPath = Join-Path $projectDir $pptxName
        if (Test-Path $pptxPath) {
            $reportExists = $true
            $reportFile = $pptxPath
            $actualFileName = $pptxName
        }

        $pptxPath = if ($reportFile) { $reportFile } else { $pptxPath }

        # Determine status and color
        $reportCreatedDate = $null
        $lagDays           = $null
        $status            = "Missing"
        $color             = "Red"

        if ($reportExists) {
            $reportCreatedDate = (Get-Item $pptxPath).LastWriteTime.Date
            $lagDays           = ($reportCreatedDate - $endDate.Date).Days
            $color             = Get-StatusAndColor $lagDays
            $status            = $color

            $icon = switch ($color) {
                "Green" { "[OK]" }
                "Yellow" { "[!]" }
                "Red" { "[X]" }
            }
            $lagDisplay = if ($lagDays -ge 0) { "+$lagDays" } else { "$lagDays" }
            Write-Host ("  {0} {1,-14} Lag: {2} days" -f $icon, $sprintNameFormatted, $lagDisplay) -ForegroundColor $(
                switch ($color) { "Green" { "Green" } "Yellow" { "Yellow" } "Red" { "Red" } }
            )
        } else {
            Write-Host ("  [X] {0,-14} (MISSING)" -f $sprintNameFormatted) -ForegroundColor Red
            $missingCount++
        }

        $entries.Add([PSCustomObject]@{
            sprintName        = $sprintNameFormatted
            sprintStartDate   = if ($startDate) { $startDate.ToString("yyyy-MM-dd") } else { $null }
            sprintEndDate     = if ($endDate) { $endDate.ToString("yyyy-MM-dd") } else { $null }
            reportCreatedDate = if ($reportCreatedDate) { $reportCreatedDate.ToString("yyyy-MM-dd") } else { $null }
            lagDays           = $lagDays
            status            = $status
            parseError        = $null
            expectedFileName  = $pptxName
            actualFileName    = $actualFileName
        })

        $sprintCount++
    }

    if ($sprintCount -eq 0 -and $parseErrorCount -eq 0 -and $missingCount -eq 0) {
        Write-Host "  No sprints found on or after $FromDate." -ForegroundColor Gray
    }
    if ($parseErrorCount -gt 0) {
        Write-Host ("  {0} sprint(s) could not be parsed." -f $parseErrorCount) -ForegroundColor Red
    }
    if ($missingCount -gt 0) {
        Write-Host ("  {0} report(s) missing." -f $missingCount) -ForegroundColor Yellow
    }
    Write-Host ""

    # Add project entry if it has any sprints
    if ($entries.Count -gt 0) {
        $sortedEntries = @($entries | Sort-Object {
            if ($_.sprintEndDate) {
                [DateTime]::Parse($_.sprintEndDate)
            } else {
                [DateTime]::MaxValue
            }
        })
        $matrix.Add([PSCustomObject]@{
            application = $projectName
            entries     = $sortedEntries
        })
    }
}

# ---------------------------------------------------------------------------
# 6. Sort and write JSON
# ---------------------------------------------------------------------------
$sorted = $matrix | Sort-Object application

$output = [PSCustomObject]@{
    generated = (Get-Date).ToString("yyyy-MM-dd")
    fromDate  = $fromDateParsed.ToString("yyyy-MM-dd")
    matrix    = @($sorted)
}

$json = $output | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($OutputJson, $json, [System.Text.Encoding]::UTF8)

# ---------------------------------------------------------------------------
# 7. Summary
# ---------------------------------------------------------------------------
$totalEntries = 0
$greenCount   = 0
$yellowCount  = 0
$redCount     = 0
$missingCount = 0

foreach ($app in $matrix) {
    foreach ($entry in $app.entries) {
        $totalEntries++
        switch ($entry.status) {
            "Green"   { $greenCount++   }
            "Yellow"  { $yellowCount++  }
            "Red"     { $redCount++     }
            "Missing" { $missingCount++ }
        }
    }
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Done" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Total entries : $totalEntries"   -ForegroundColor White
Write-Host "  Green (0 day lag)  : $greenCount"   -ForegroundColor Green
Write-Host "  Yellow (1-2 days)  : $yellowCount"  -ForegroundColor Yellow
Write-Host "  Red (3+ days)      : $redCount"     -ForegroundColor Red
Write-Host "  Missing            : $missingCount" -ForegroundColor $(if ($missingCount -gt 0) { "Red" } else { "White" })
Write-Host ""
Write-Host "  Output: $OutputJson" -ForegroundColor Cyan
Write-Host ""
