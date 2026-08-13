# Sprint Reports Metrics Matrix - Implementation Summary

## ✅ Completed Implementation

All files have been created and integrated. Here's what was delivered:

---

## 📁 Files Created

### 1. PowerShell Script
- **File:** `generate-sprint-reports-availability.ps1`
- **Location:** `C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports\`
- **Purpose:** Generates the metrics JSON file with customizable input/output paths
- **Parameters:** `-AppSettingsPath`, `-OutputJson`, `-FromDate`

### 2. C# Models
- **File:** `GenerateDeliveryReports.Models\SprintReportsAvailabilityMatrix.cs`
- **Classes:**
  - `SprintReportsAvailabilityMatrix` (root)
  - `ProductReportMatrix` (product/app)
  - `SprintReportEntry` (individual sprint)

### 3. Data Service
- **File:** `GenerateDeliveryReports.Data\Services\SprintReportsAvailabilityService.cs`
- **Purpose:** Loads JSON and deserializes into C# models

### 4. Blazor Component
- **File:** `GenerateDeliveryReports\Components\Pages\SprintReportMetrics.razor`
- **Route:** `/sprint-report-metrics`
- **Features:**
  - Interactive table with product grouping
  - Color-coded status badges (Green/Yellow/Red/Missing)
  - Lag days display (+0, +2, +4, etc.)
  - Summary statistics
  - Parse error section
  - Refresh button
  - Responsive design

### 5. Service Registration
- **File:** `Program.cs` (modified)
- **Change:** Added `builder.Services.AddScoped<SprintReportsAvailabilityService>();`

### 6. Navigation Menu
- **File:** `Components\Layout\NavMenu.razor` (modified)
- **Change:** Added menu item "Sprint Report Metrics" → `/sprint-report-metrics`

### 7. Documentation
- **File:** `SPRINT_REPORTS_METRICS_GUIDE.md` - Complete usage guide
- **File:** `IMPLEMENTATION_SUMMARY.md` - This file

---

## 🎯 Key Features

### Color Coding System

| Status | Lag Days | Color  | Interpretation |
|--------|----------|--------|-----------------|
| Green  | 0 days   | 🟢     | Same day delivery (excellent) |
| Yellow | 1-2 days | 🟡     | Minor delay (acceptable) |
| Red    | 3+ days  | 🔴     | Late delivery (needs attention) |
| Missing| N/A      | ⚫     | Report not found (blocker) |

### Data Included in JSON

For each sprint, the matrix captures:
- ✓ Sprint name (e.g., "Sprint 12")
- ✓ Sprint start date (ISO format)
- ✓ Sprint end date (ISO format)
- ✓ Report creation date (ISO format)
- ✓ Lag days (calculated: `reportDate - sprintEndDate`)
- ✓ Status (Green/Yellow/Red/Missing)
- ✓ Parse errors (if sprint name couldn't be parsed)

---

## 📊 JSON Schema Example

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
        }
      ]
    }
  ]
}
```

---

## 🚀 Getting Started

### Quick Start (3 steps)

1. **Generate the JSON file:**
   ```powershell
   cd C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports
   .\generate-sprint-reports-availability.ps1
   ```

2. **Place the file in the right location:**
   - Check `appsettings.json` for `SprintReportsAvailabilityJSONFileName`
   - Copy the generated JSON to that location

3. **View in the app:**
   - Navigate to `/sprint-report-metrics` or use the menu
   - Data will display in a formatted matrix table

---

## ⚙️ Script Usage Examples

### Default Run
```powershell
.\generate-sprint-reports-availability.ps1
```
- Reads: `./GenerateDeliveryReports/appsettings.json`
- Outputs: `./SprintReportsAvailability.json`
- Processes: Sprints from 2026-04-01 onward

### Custom Output Location
```powershell
.\generate-sprint-reports-availability.ps1 `
    -OutputJson "C:\Deployments\MetricsData\SprintReportsAvailability.json"
```

### Custom App Settings
```powershell
.\generate-sprint-reports-availability.ps1 `
    -AppSettingsPath "C:\Production\appsettings.json" `
    -OutputJson "C:\Deployments\SprintReportsAvailability.json"
```

### Filter by Date
```powershell
.\generate-sprint-reports-availability.ps1 `
    -FromDate "2026-07-01"  # Only July 2026 onwards
```

---

## 📋 Configuration in appsettings.json

Already configured (no changes needed):

```json
"AppSettings": {
  "SprintReportsAvailabilityJSONFileName": "SprintReportsAvailability.json",
  "OneDriveLocation": "C:\\Users\\coolg\\OneDrive - Relevantz\\",
  "ReportAndDataFolder": "Projects",
  ...
}
```

To customize the JSON file location, just update `SprintReportsAvailabilityJSONFileName`.

---

## 🔄 Automation Options

### Option A: Windows Task Scheduler
Create a scheduled task to run the script daily/weekly:

```powershell
# Run this once to create the scheduled task
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
  -Argument "-NoProfile -ExecutionPolicy Bypass -File `"C:\path\to\generate-sprint-reports-availability.ps1`""

$trigger = New-ScheduledTaskTrigger -Daily -At "08:00 AM"

Register-ScheduledTask -TaskName "GenerateSprintMetrics" `
  -Action $action -Trigger $trigger -RunLevel Highest
```

### Option B: Existing Worker App
Add the script to your `GenerateDeliveryReports.Worker` console app if you want it to run as part of your existing automation.

### Option C: Manual Run
Keep it simple - run the script whenever you need updated metrics.

---

## ✨ Visual Features in the App

### Page Components

1. **Header** - Title + Refresh button
2. **Metadata** - Generation date, from date, last modified time
3. **Legend** - Color and lag day ranges explained
4. **Summary Stats** - Counts of Green/Yellow/Red/Missing
5. **Main Table** - Sortable, responsive matrix with:
   - Product names (grouped in rows)
   - Sprint details (name, dates)
   - Report status (color + lag days)
6. **Error Section** - Lists unparseable sprint names

---

## 🔍 Monitoring & Insights

Use this dashboard to:

- ✓ **Track SLAs** - Monitor report delivery timeliness
- ✓ **Identify Trends** - See which products consistently miss deadlines
- ✓ **Spot Issues** - Quickly find missing reports
- ✓ **Measure Improvement** - Track over-time progress
- ✓ **Team Performance** - Identify high/low performers

---

## ⚠️ Important Notes

1. **Sprint Name Format** - Excel must use: `Sprint X (DD-MMM-YYYY to DD-MMM-YYYY)`
   - Example: `Sprint 12 (01-Apr-2026 to 14-Apr-2026)`

2. **Report File Naming** - Must follow pattern:
   - `GlobalPayments-[ProjectName]-DeliveryQualitySummaryReport-[SprintName].pptx`

3. **OneDrive Path** - Script requires access to the configured OneDrive location

4. **ImportExcel Module** - Required for the PowerShell script:
   ```powershell
   Install-Module ImportExcel -Scope CurrentUser -Force
   ```

---

## 📞 Troubleshooting

### Problem: "No metrics file found"
**Solution:** Run the script to generate the JSON file first.

### Problem: Parse errors appear
**Solution:** Update sprint names in Excel to match the expected format.

### Problem: Reports show as missing
**Solution:** Verify report files exist at the expected location and have correct names.

### Problem: Script fails with module error
**Solution:** Install ImportExcel module with:
```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

---

## 🎓 Next Steps

1. ✅ Run the PowerShell script to generate the first metrics JSON
2. ✅ Place the JSON file in the configured location
3. ✅ Open the app and navigate to Sprint Report Metrics
4. ✅ Verify all products and sprints display correctly
5. ✅ Set up automation (if desired) to update metrics regularly
6. ✅ Share the dashboard with your team for tracking delivery performance

---

## 📝 Summary

You now have a complete **Sprint Report Delivery Metrics Matrix** system that:

- ✅ Reads data from existing Excel workbooks
- ✅ Calculates lag days automatically
- ✅ Color-codes delivery timeliness
- ✅ Displays results in an interactive Blazor component
- ✅ Supports customizable output locations
- ✅ Integrates seamlessly with your existing app

The system is ready to use. Just run the script and navigate to the new page!
