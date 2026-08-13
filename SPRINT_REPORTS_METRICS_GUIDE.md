# Sprint Reports Metrics Matrix - Implementation Guide

## Overview

This feature adds a **Sprint Report Delivery Metrics Matrix** to the app that tracks and visualizes report delivery timeliness across all products. The matrix shows:

- **Products/Applications** (rows)
- **Sprints** with their delivery status and lag days (columns)
- **Color coding** indicating delivery speed
- **Missing reports** highlighted for tracking

## Components Created

### 1. PowerShell Script: `generate-sprint-reports-availability.ps1`

**Location:** `C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports\generate-sprint-reports-availability.ps1`

**Purpose:** Generates the JSON data file for the metrics matrix by reading Excel data from all projects.

**Parameters:**

```powershell
-AppSettingsPath     # Path to appsettings.json (default: auto-detected)
-OutputJson          # Where to save the JSON file (default: ./SprintReportsAvailability.json)
-FromDate            # Cutoff date, skip earlier sprints (default: "2026-04-01")
```

**Usage Examples:**

```powershell
# Default: uses app settings from local project, saves to current directory
.\generate-sprint-reports-availability.ps1

# Custom output location
.\generate-sprint-reports-availability.ps1 -OutputJson "C:\Deployments\SprintReportsAvailability.json"

# Custom app settings path
.\generate-sprint-reports-availability.ps1 -AppSettingsPath "C:\path\to\appsettings.json"

# Only process sprints ending on or after a specific date
.\generate-sprint-reports-availability.ps1 -FromDate "2026-06-01"

# Combine all parameters
.\generate-sprint-reports-availability.ps1 `
    -AppSettingsPath "C:\path\to\appsettings.json" `
    -OutputJson "C:\Deployments\SprintReportsAvailability.json" `
    -FromDate "2026-07-01"
```

**Output Metrics Logged to Console:**

- ✓ Green (0 days lag): Reports created same day as sprint end
- ! Yellow (1-2 days): Reports created 1-2 days after sprint end
- ✕ Red (3+ days): Reports created 3+ days after sprint end
- Missing: No report found

### 2. Data Models

**File:** `GenerateDeliveryReports.Models\SprintReportsAvailabilityMatrix.cs`

**Classes:**

- `SprintReportsAvailabilityMatrix` - Root object containing generated date, fromDate, and matrix data
- `ProductReportMatrix` - Product/Application with list of sprint entries
- `SprintReportEntry` - Individual sprint with dates, lag days, and status

### 3. Data Service

**File:** `GenerateDeliveryReports.Data\Services\SprintReportsAvailabilityService.cs`

**Purpose:** Loads and deserializes the JSON file into C# objects for the Blazor component.

**Key Method:**
```csharp
public (SprintReportsAvailabilityMatrix? Matrix, DateTime? LastModified, string? Error) Load()
```

### 4. Blazor Component

**File:** `GenerateDeliveryReports\Components\Pages\SprintReportMetrics.razor`

**Route:** `/sprint-report-metrics`

**Features:**

- Displays all products and their sprints in a responsive table
- Color-coded status badges:
  - **Green** = 0 days lag (same day)
  - **Yellow** = 1-2 days lag
  - **Red** = 3+ days lag
  - **Missing** = No report found
- Shows lag days (+0, +2, +4, etc.)
- Lists parse errors if sprint names couldn't be parsed
- Displays file generation timestamp and metadata
- Summary statistics showing count of each status
- Refresh button to reload the data

### 5. Service Registration

**File:** `Program.cs`

Added service registration:
```csharp
builder.Services.AddScoped<SprintReportsAvailabilityService>();
```

### 6. Navigation

**File:** `Components\Layout\NavMenu.razor`

Added menu link: "Sprint Report Metrics" → `/sprint-report-metrics`

---

## JSON File Schema

**File Name:** `SprintReportsAvailability.json` (configurable in `appsettings.json`)

**Location:** Configured via `SprintReportsAvailabilityJSONFileName` setting

**Structure:**

```json
{
  "generated": "2026-08-11",
  "fromDate": "2026-04-01",
  "matrix": [
    {
      "application": "Hosted Payments",
      "entries": [
        {
          "sprintName": "Sprint 12",
          "sprintStartDate": "2026-04-01",
          "sprintEndDate": "2026-04-14",
          "reportCreatedDate": "2026-04-14",
          "lagDays": 0,
          "status": "Green",
          "parseError": null
        },
        {
          "sprintName": "Sprint 13",
          "sprintStartDate": "2026-04-15",
          "sprintEndDate": "2026-04-28",
          "reportCreatedDate": "2026-04-30",
          "lagDays": 2,
          "status": "Yellow",
          "parseError": null
        }
      ]
    }
  ]
}
```

---

## Configuration

### In `appsettings.json`

The following setting already exists:

```json
"AppSettings": {
  "SprintReportsAvailabilityJSONFileName": "SprintReportsAvailability.json",
  ...
}
```

**To customize the location:**

```json
"SprintReportsAvailabilityJSONFileName": "C:\\Deployments\\Global\\SprintReportsAvailability.json"
```

---

## Workflow

### Step 1: Generate the JSON File

Run the PowerShell script to generate the data:

```powershell
cd C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports
.\generate-sprint-reports-availability.ps1
```

This creates `SprintReportsAvailability.json` in the current directory (or custom location via `-OutputJson`).

### Step 2: Place the JSON File

Copy the generated JSON file to the location specified in `appsettings.json`:

```powershell
# If using default config, it looks for the file here:
Copy-Item .\SprintReportsAvailability.json $Env:APPDATA\..\Local\Temp\...\SprintReportsAvailability.json

# Or in the app's wwwroot directory:
Copy-Item .\SprintReportsAvailability.json .\GenerateDeliveryReports\wwwroot\SprintReportsAvailability.json
```

### Step 3: View in the App

Navigate to **Sprint Report Metrics** from the left navigation menu or visit `/sprint-report-metrics`.

---

## Automation (Optional)

To auto-generate the metrics on a schedule:

### Using Windows Task Scheduler

Create a scheduled task to run the script daily/weekly:

```powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
  -Argument "-NoProfile -ExecutionPolicy Bypass -File `"C:\path\to\generate-sprint-reports-availability.ps1`" -OutputJson `"C:\Deployments\SprintReportsAvailability.json`""

$trigger = New-ScheduledTaskTrigger -Daily -At "09:00 AM"

Register-ScheduledTask -TaskName "GenerateSprintMetrics" -Action $action -Trigger $trigger -RunLevel Highest
```

### Using the App's Worker

If you want to integrate this into the existing `GenerateDeliveryReports.Worker` console app, you can add a similar service and job runner.

---

## Troubleshooting

### "No metrics file found" Message

**Cause:** The JSON file hasn't been generated or is in the wrong location.

**Solution:**
1. Run the PowerShell script: `.\generate-sprint-reports-availability.ps1`
2. Verify the file location matches `SprintReportsAvailabilityJSONFileName` in `appsettings.json`
3. Check file permissions (the app must have read access)

### Parse Errors Show Up

**Cause:** Sprint names in Excel don't match the expected format.

**Expected Format:** `Sprint X (DD-MMM-YYYY to DD-MMM-YYYY)`

**Examples:**
- ✓ `Sprint 12 (01-Apr-2026 to 14-Apr-2026)`
- ✓ `Sprint 13 (15-Apr-2026 to 28-Apr-2026)`
- ✗ `Sprint 12` (missing dates)
- ✗ `Sprint 12 (Apr 1 to Apr 14)` (wrong format)

**Solution:** Update the sprint names in the Excel data sheets to include the date range.

### Reports Show as "Missing"

**Cause:** The PPTX file wasn't found at the expected location.

**Expected Path Pattern:** `[ProjectFolder]/GlobalPayments-[ProjectName]-DeliveryQualitySummaryReport-[SprintName].pptx`

**Solution:**
1. Verify reports are being generated
2. Check report file naming matches the pattern
3. Ensure the OneDriveLocation path in `appsettings.json` is correct

---

## Data Interpretation

### Status Colors

| Color  | Meaning | Lag Days |
|--------|---------|----------|
| Green  | On time | 0 days   |
| Yellow | Slightly late | 1-2 days |
| Red    | Late   | 3+ days  |
| Missing | Not found | N/A |

### Metrics to Monitor

- **Green Percentage:** What % of reports are delivered on the same day?
- **Red Percentage:** What % have >3 day delays?
- **Missing:** Are there products consistently missing reports?
- **Trends:** Has on-time delivery improved over time?

---

## Next Steps

1. **Run the script** to generate the initial JSON file
2. **Test the page** by viewing `/sprint-report-metrics`
3. **Schedule the script** to run daily/weekly for ongoing data updates
4. **Monitor trends** to identify products needing process improvements
5. **Share with stakeholders** to track delivery SLAs

---

## Support

For issues or modifications:

1. Check the troubleshooting section above
2. Review the script console output for specific errors
3. Verify Excel data format in the "Data" worksheet of each project
4. Check `appsettings.json` configuration
