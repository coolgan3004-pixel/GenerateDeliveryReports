# =============================================================================
# generate-sprint-status.ps1
# =============================================================================
# Reads every project's DataFileName Excel workbook, extracts sprint names and
# dates, checks whether the corresponding PPTX delivery report exists, and
# writes a JSON file that the Blazor app (or anyone on the team) can consume.
#
# Requirements:
#   ImportExcel module -- install once per machine:
#   Install-Module ImportExcel -Scope CurrentUser -Force
#
# Usage:
#   .\generate-sprint-status.ps1
#   .\generate-sprint-status.ps1 -AppSettingsPath "C:\path\to\appsettings.json"
#   .\generate-sprint-status.ps1 -OutputJson "C:\Deployments\sprint-report-status.json"
#   .\generate-sprint-status.ps1 -FromDate "2026-06-01"   # override the cutoff date
# =============================================================================

param(
    # Path to the source appsettings.json (dev machine copy, not the published one)
    [string]$AppSettingsPath = $null,

    # Where to write the output JSON file
    [string]$OutputJson = $null,

    # Ignore sprints whose end date is before this date (default: Apr 1 2026,
    # matching the ReportWorker cutoff -- earlier sprints may not follow the naming format)
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
    $OutputJson = Join-Path $scriptDir "sprint-report-status.json"
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Sprint Report Status -- JSON Generator"   -ForegroundColor Cyan
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
        "d-MMM-yyyy","dd-MMM-yyyy",
        "MMM-d-yyyy","MMM-dd-yyyy",
        "d-MMMM-yyyy","dd-MMMM-yyyy",
        "dd/MM/yyyy","d/M/yyyy","MM/dd/yyyy","M/d/yyyy",
        "dd-MM-yyyy","d-M-yyyy","dd.MM.yyyy"
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

# ---------------------------------------------------------------------------
# 5. Process each project
# ---------------------------------------------------------------------------
$entries = [System.Collections.Generic.List[object]]::new()

foreach ($project in $projects) {
    $projectName  = $project.ProjectName
    $dataFilePath = Join-Path $oneDriveBase (Join-Path $dataFolder $project.DataFileName)
    $projectDir   = Split-Path $dataFilePath -Parent

    Write-Host "[$projectName]" -ForegroundColor White
    Write-Host "  DataFile: $dataFilePath" -ForegroundColor Gray

    # Find the actual file (handles OneDrive duplicate "File (1).xlsx" copies)
    $actualFile = Get-RecentSimilarFile $dataFilePath
    if (-not $actualFile) {
        Write-Host "  SKIP -- file not found." -ForegroundColor Yellow
        continue
    }
    if ($actualFile.FullName -ne $dataFilePath) {
        Write-Host "  Using   : $($actualFile.FullName)" -ForegroundColor Gray
    }

    # Read the Data worksheet (row 1 = headers, same as C# ReadSpecificColumnsFromRange)
    try {
        $rows = Import-Excel -Path $actualFile.FullName -WorksheetName "Data" -ErrorAction Stop
    }
    catch {
        Write-Host "  SKIP -- could not open Data sheet: $_" -ForegroundColor Yellow
        continue
    }

    $sprintCount      = 0
    $parseErrorCount  = 0
    foreach ($row in $rows) {
        $sprintName = "$($row.Sprint)".Trim()
        if (-not $sprintName -or $sprintName -eq '' -or $sprintName -eq 'nan') { continue }

        # ---- Attempt to parse start / end dates from the sprint name ----
        # Expected format: "Sprint X (DD-MMM-YYYY to DD-MMM-YYYY)"
        $startDate  = $null
        $endDate    = $null
        $parseError = $null

        if ($sprintName -match '\((\d{1,2}[\/\-\.]\w{1,9}[\/\-\.]\d{2,4})\s*to\s*(\d{1,2}[\/\-\.]\w{1,9}[\/\-\.]\d{2,4})\)') {
            $startDate = Parse-SprintDate $Matches[1]
            $endDate   = Parse-SprintDate $Matches[2]
            if (-not $startDate) { $parseError = "Unrecognised start date format: '$($Matches[1])'" }
            if (-not $endDate)   { $parseError = if ($parseError) { "$parseError; unrecognised end date format: '$($Matches[2])'" }
                                                 else              { "Unrecognised end date format: '$($Matches[2])'" } }
        } else {
            $parseError = "Sprint name does not contain a recognisable date range: '$sprintName'"
        }

        # ---- Parse error: log to console and record in JSON ----
        if ($parseError) {
            Write-Host ("  [??] {0}" -f $sprintName) -ForegroundColor Red
            Write-Host ("       PARSE ERROR: {0}" -f $parseError) -ForegroundColor Red
            $entries.Add([PSCustomObject]@{
                month             = $null
                application       = $projectName
                sprintName        = $sprintName
                sprintStartDate   = $null
                sprintEndDate     = $null
                reportStatus      = "ParseError"
                reportCreatedDate = $null
                daysDiff          = $null
                parseError        = $parseError
            })
            $parseErrorCount++
            continue
        }

        # ---- Intentional cutoff: silently skip sprints before FromDate ----
        # (Earlier sprints may not follow the naming convention -- not an error)
        if ($endDate.Date -lt $fromDateParsed.Date) { continue }

        # Strip the date part to get the name used in the PPTX filename
        # "Sprint 12 (01-Apr-2026 to 14-Apr-2026)" -> "Sprint 12"
        $sprintNameFormatted = if ($sprintName.Contains('(')) {
            $sprintName.Substring(0, $sprintName.IndexOf('(')).Trim()
        } else {
            $sprintName.Trim()
        }

        # Build expected PPTX path (same convention as DataProcessor.cs lines 142-144)
        $pptxName = "GlobalPayments-$projectName-DeliveryQualitySummaryReport-$sprintNameFormatted.pptx"
        $pptxPath = Join-Path $projectDir $pptxName

        $reportExists = Test-Path $pptxPath

        # Determine status
        $status = if ($endDate.Date -gt [DateTime]::Today) {
            "Upcoming"
        } elseif ($reportExists) {
            "Available"
        } else {
            "Missing"
        }

        # Days between sprint end and the report being written
        $reportCreatedDate = $null
        $daysDiff          = $null
        if ($reportExists) {
            $reportCreatedDate = (Get-Item $pptxPath).LastWriteTime.Date
            $daysDiff          = ($reportCreatedDate - $endDate.Date).Days
        }

        $entries.Add([PSCustomObject]@{
            month             = $endDate.ToString("MMMM yyyy")
            application       = $projectName
            sprintName        = $sprintNameFormatted
            sprintStartDate   = $startDate.ToString("yyyy-MM-dd")
            sprintEndDate     = $endDate.ToString("yyyy-MM-dd")
            reportStatus      = $status
            reportCreatedDate = if ($reportCreatedDate) { $reportCreatedDate.ToString("yyyy-MM-dd") } else { $null }
            daysDiff          = $daysDiff
            parseError        = $null
        })

        $icon = switch ($status) { "Available" { "OK" } "Missing" { "!!" } "Upcoming" { "--" } }
        Write-Host ("  [{0}] {1,-14}  {2}" -f $icon, $sprintNameFormatted, $status) -ForegroundColor $(
            switch ($status) { "Available" { "DarkGray" } "Missing" { "Yellow" } "Upcoming" { "Gray" } }
        )
        $sprintCount++
    }

    if ($sprintCount -eq 0 -and $parseErrorCount -eq 0) {
        Write-Host "  No sprints found on or after $FromDate." -ForegroundColor Gray
    }
    if ($parseErrorCount -gt 0) {
        Write-Host ("  {0} sprint(s) could not be parsed -- check sprint name format in the Data sheet." -f $parseErrorCount) -ForegroundColor Red
    }
    Write-Host ""
}

# ---------------------------------------------------------------------------
# 6. Sort and write JSON
# ---------------------------------------------------------------------------
$sorted = $entries | Sort-Object {
    if ($_.sprintEndDate) { [DateTime]::Parse($_.sprintEndDate) } else { [DateTime]::MaxValue }
}, application

$output = [PSCustomObject]@{
    generated = (Get-Date).ToString("yyyy-MM-dd")
    fromDate  = $fromDateParsed.ToString("yyyy-MM-dd")
    entries   = @($sorted)
}

$json = $output | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($OutputJson, $json, [System.Text.Encoding]::UTF8)

# ---------------------------------------------------------------------------
# 7. Summary
# ---------------------------------------------------------------------------
$total       = $entries.Count
$available   = ($entries | Where-Object { $_.reportStatus -eq "Available"  }).Count
$missing     = ($entries | Where-Object { $_.reportStatus -eq "Missing"    }).Count
$upcoming    = ($entries | Where-Object { $_.reportStatus -eq "Upcoming"   }).Count
$parseErrors = ($entries | Where-Object { $_.reportStatus -eq "ParseError" }).Count

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Done" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Total entries : $total"       -ForegroundColor White
Write-Host "  Available     : $available"   -ForegroundColor Green
Write-Host "  Missing       : $missing"     -ForegroundColor $(if ($missing     -gt 0) { "Yellow" } else { "White" })
Write-Host "  Upcoming      : $upcoming"    -ForegroundColor Gray
Write-Host "  Parse errors  : $parseErrors" -ForegroundColor $(if ($parseErrors -gt 0) { "Red"    } else { "White" })
Write-Host ""
Write-Host "  Output: $OutputJson" -ForegroundColor Cyan
Write-Host ""
