# Sprint Report Metrics - Quick Reference

## 🚀 Quick Start

### Generate Metrics JSON
```powershell
cd C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports
.\generate-sprint-reports-availability.ps1
```

### View in App
Navigate to: **Sprint Report Metrics** (from left menu or `/sprint-report-metrics`)

---

## 🔧 Script Commands

### Standard (default paths & settings)
```powershell
.\generate-sprint-reports-availability.ps1
```

### Custom output location
```powershell
.\generate-sprint-reports-availability.ps1 -OutputJson "C:\Deployments\SprintReportsAvailability.json"
```

### Custom app settings & output
```powershell
.\generate-sprint-reports-availability.ps1 `
    -AppSettingsPath "C:\path\to\appsettings.json" `
    -OutputJson "C:\Deployments\SprintReportsAvailability.json"
```

### Only recent sprints (after July 1, 2026)
```powershell
.\generate-sprint-reports-availability.ps1 -FromDate "2026-07-01"
```

---

## 🎨 Color Key

| Color | Lag Days | Meaning |
|-------|----------|---------|
| 🟢 Green | 0 | Same day delivery |
| 🟡 Yellow | 1-2 | Slightly late |
| 🔴 Red | 3+ | Late delivery |
| ⚫ Missing | — | Report not found |

---

## 📍 Key File Locations

| Component | Path |
|-----------|------|
| **Script** | `C:\Repository\GenerateDeliveryReports\GenerateDeliveryReports\generate-sprint-reports-availability.ps1` |
| **Blazor Component** | `GenerateDeliveryReports\Components\Pages\SprintReportMetrics.razor` |
| **Data Models** | `GenerateDeliveryReports.Models\SprintReportsAvailabilityMatrix.cs` |
| **Service** | `GenerateDeliveryReports.Data\Services\SprintReportsAvailabilityService.cs` |
| **JSON Output** | See `SprintReportsAvailabilityJSONFileName` in `appsettings.json` |

---

## ⚙️ Configuration

**File:** `appsettings.json`

```json
"AppSettings": {
  "SprintReportsAvailabilityJSONFileName": "SprintReportsAvailability.json",
  "OneDriveLocation": "C:\\Users\\coolg\\OneDrive - Relevantz\\",
  "ReportAndDataFolder": "Projects",
  ...
}
```

---

## 🎯 Expected Sprint Name Format

### In Excel (Data worksheet):
```
Sprint 12 (01-Apr-2026 to 14-Apr-2026)
Sprint 13 (15-Apr-2026 to 28-Apr-2026)
```

### Report Filename Pattern:
```
GlobalPayments-Hosted Payments-DeliveryQualitySummaryReport-Sprint 12.pptx
```

---

## ✅ Verification Checklist

- [ ] Script runs without errors
- [ ] JSON file is created at the configured location
- [ ] JSON file appears in `/sprint-report-metrics` page
- [ ] Products and sprints display correctly
- [ ] Colors are properly assigned (Green/Yellow/Red)
- [ ] Lag days are calculated correctly
- [ ] Missing reports are identified

---

## 📊 JSON Schema (Quick View)

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
          "status": "Green"
        }
      ]
    }
  ]
}
```

---

## 🛠️ Common Tasks

### Regenerate metrics
```powershell
.\generate-sprint-reports-availability.ps1 -OutputJson "C:\path\to\SprintReportsAvailability.json"
```

### Schedule daily run (Windows Task Scheduler)
```powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
  -Argument "-NoProfile -ExecutionPolicy Bypass -File `"C:\...\generate-sprint-reports-availability.ps1`""
$trigger = New-ScheduledTaskTrigger -Daily -At "08:00 AM"
Register-ScheduledTask -TaskName "GenerateSprintMetrics" -Action $action -Trigger $trigger -RunLevel Highest
```

### Move JSON file to new location
```powershell
Copy-Item ".\SprintReportsAvailability.json" "C:\Deployments\SprintReportsAvailability.json"
# Then update SprintReportsAvailabilityJSONFileName in appsettings.json
```

---

## ❌ Troubleshooting

| Issue | Solution |
|-------|----------|
| "ImportExcel module not found" | `Install-Module ImportExcel -Scope CurrentUser -Force` |
| "No metrics file found" in app | Run script to generate JSON, verify path in appsettings.json |
| Parse errors appear | Check sprint name format in Excel matches `Sprint X (DD-MMM-YYYY to DD-MMM-YYYY)` |
| Reports show as missing | Verify PPTX files exist with correct naming pattern |
| Can't find generated file | Check script console output for actual output path |

---

## 📚 Full Documentation

For detailed info, see:
- **Setup & Configuration:** `SPRINT_REPORTS_METRICS_GUIDE.md`
- **Implementation Details:** `IMPLEMENTATION_SUMMARY.md`

---

## 🔗 Related Features

- **Sprint Report Status** - View status by month (`/sprint-report-status`)
- **Metrics Trends** - Historical trends (`/metrics-trends`)
- **Sprint Dashboard** - Overall dashboard (`/sprint-dashboard`)

---

**Version:** 1.0  
**Created:** 2026-08-11  
**Last Updated:** 2026-08-11
