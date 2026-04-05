# Plan: Sprint Report Generation — Blazor Server Web Application (.NET 8)

## Problem Statement

Rebuild the Sprint Report Generation feature (originally Tab 3 of a WinForms desktop app) as a **Blazor Server web application** targeting **.NET 8**. The app reads sprint metrics from Excel sheets, computes a quality score, generates a chart, lets the user review/edit narrative sections (Sprint Summary, Highlights, Retrospective), and produces a formatted `.pptx`/`.pdf` report — all from a browser UI.

This plan serves as a **rebuild prompt** — complete enough to generate the entire application from scratch as a Blazor Server app.

> **Why Blazor Server (not Blazor WASM)?**
> The app must read/write Excel files, export charts and generate PowerPoint/PDF files on disk. Blazor Server runs on the server (SignalR-connected), so all file I/O and NuGet library calls work directly — no API layer needed. Blazor WASM runs in the browser and cannot access the server file system.

> **Note on ".NET Standard 8":** There is no ".NET Standard 8". The target is **.NET 8** (the modern unified platform). Use `<TargetFramework>net8.0</TargetFramework>`.

---

## Architecture Overview

Clean layered architecture across 4 projects in one solution:

| Layer | Project | Responsibility |
|---|---|---|
| UI | `SprintReport.Web` (Blazor Server) | Razor components, pages, forms |
| Service | `SprintReport.Services` | Business logic, orchestration (replaces Presenter) |
| Data | `SprintReport.Data` | Excel read/write, chart export, PPT/PDF generation |
| Model | `SprintReport.Model` | Shared POCOs/DTOs, interfaces |

**Target framework:** `.NET 8` — `net8.0` for all projects

**Key NuGet packages:**
- `EPPlus` (v7+) — Excel read/write (`ExcelPackage.License.SetNonCommercialPersonal(...)`)
- `Spire.Xls` — Excel chart-to-PNG image export
- `Spire.Presentation` — PowerPoint generation from `.pptx` template
- `Newtonsoft.Json` — `appsettings.json` / `appConfig.json` deserialization
- `Serilog.AspNetCore` — structured file logging
- `Microsoft.AspNetCore.Components.Web` — Blazor component model (included with SDK)
- `MudBlazor` or `Bootstrap` — UI component styling (choose one; MudBlazor recommended for data grids and forms)

---

---

## Solution Structure

```
SprintReport.sln
├── SprintReport.Web/              ← Blazor Server app (net8.0)
│   ├── Components/
│   │   ├── App.razor
│   │   ├── Routes.razor
│   │   └── Pages/
│   │       └── SprintReport.razor     ← Main page (the 3-step wizard)
│   ├── Layout/
│   │   ├── MainLayout.razor
│   │   └── NavMenu.razor
│   ├── wwwroot/                   ← Static assets
│   ├── Program.cs                 ← DI registration, Serilog setup
│   └── appsettings.json           ← App configuration
├── SprintReport.Services/         ← Business logic (net8.0)
│   └── SprintReportService.cs
├── SprintReport.Data/             ← Data access (net8.0)
│   ├── Config/AppConfiguration.cs
│   ├── Concrete/DataProcessor.cs
│   ├── Concrete/ExcelWrapper.cs
│   └── Extensions/
├── SprintReport.Model/            ← Shared models (net8.0)
│   ├── SprintMetrics.cs
│   ├── PPTReportData.cs
│   ├── ReportDataParameters.cs
│   ├── AppSettings.cs
│   ├── Project.cs
│   └── DashboardData.cs
```

---

## Configuration (`appsettings.json`)

All paths and project metadata are stored in `appsettings.json` under an `AppSettings` section and bound via `IOptions<AppSettings>` in `Program.cs`:

```json
{
  "AppSettings": {
    "UseDownloadsFolder": "true",
    "DownloadsFolder": "C:\\...\\Downloads",
    "OneDriveLocation": "",
    "ReportAndDataFolder": "",
    "MetricsFolder": "",
    "SprintMetricsReportTemplatePath": "Templates\\GlobalPayments-DeliveryQualitySummaryReport_Template.pptx",
    "PMOEmailContent": "Hi Team,\n\nPlease find the report ## at #Link#.\n\nRegards",
    "Projects": [
      {
        "ProjectName": "ProjectA",
        "DataFileName": "ProjectA_Data.xlsx",
        "MetricsSheetPath": ["ProjectA - Sprint Metrics.xlsx"],
        "ProjectFolderOneDriveLink": "https://..."
      }
    ],
    "EmailSettings": {
      "UserName": "",
      "Password": "",
      "FromEmailAddress": ""
    }
  }
}
```

In `Program.cs`:
```csharp
builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));
builder.Services.AddSingleton<IDataProcessor, DataProcessor>();
builder.Services.AddScoped<SprintReportService>();
builder.Host.UseSerilog((ctx, cfg) => cfg.WriteTo.File("LogFiles/log.txt", rollingInterval: RollingInterval.Day));
```

`AppSettings` has a computed `TempPath` = `wwwroot/downloads` (auto-created, served as static files for download links).

---

## Data Model

### `SprintMetrics` (Model)
Represents one sprint row in the `Data` sheet of the project Excel file:

| Property | Excel Column Header | Type |
|---|---|---|
| Sprint | Sprint | string |
| Committed | Committed | object (long) |
| Delivered | Delivered | object (long) |
| CommitmentIndex | Commitment Index | object (double, formula-driven) |
| Velocity | Velocity | object (double) |
| CodeQualityIndex | Code Coverage | object (string, e.g. "85%") |
| CodeReviewCommentsInternal | Code Review Comments - Internal | object (int) |
| CodeReviewCommentsExternal | Code Review Comments - External | object (int) |
| QADefects | QA Defects | object (int) |
| EscapedDefects | Escaped Defects | object (int) |
| BacklogHealth | Backlog Health | object (int) |
| LastSprint | Last Sprint? | string ("Yes"/"No") |
| Remarks | Remarks | object (string) |

### `PPTReportData` (Model)
Returned by `GenerateChart`:
- `ImagePath` — path to the exported chart PNG
- `Score` — score string read from Scorecard cell B3 (e.g. "0.92")
- `SprintSummary` — string array (lines)
- `SprintHighlights` — string array (lines)
- `SprintRetrospective` — string array (lines, structured with categories)

### `ReportDataParameters` (Model)
Passed into `GeneratePresentation`:
- `ProjectName`, `SprintNameWithDate` (e.g. `"Sprint 10 (01-Jan-2025 to 14-Jan-2025)"`)
- `SprintName` — computed: substring before `"("`
- `ImagePath`, `SprintScore`, `SprintSummary[]`, `SprintHighlights[]`, `SprintRetrospective[]`

---

## Excel File Structure

### Project Data Excel (`DataFileName`)
Used as the **master scorecard** per project.

| Sheet | Purpose |
|---|---|
| `Data` | Sprint metrics table. Headers in row 1: Sprint, Committed, Delivered, Commitment Index, Velocity, Code Coverage, Code Review Comments - Internal, Code Review Comments - External, QA Defects, Escaped Defects, Backlog Health, Last Sprint?, Remarks |
| `Scorecard` | Chart + score. Cell `B5` = sprint name. Cell `B3` = calculated score (formula). Range `B8:B13` = recalculated metrics. Contains 1 chart (index 0). |

### Sprint Metrics Excel (`MetricsSheetPath[n]`)
Per-team metrics file.

| Sheet | Purpose |
|---|---|
| `Dashboard` | Headers in row 2: Sprint #, Sprint Start Date, Sprint End Date, Assigned, Completed, Remarks. Data starts row 3. |

---

## Excel Data Reading (`ExcelWrapper`)

`ExcelWrapper` uses **EPPlus** and provides:
- `ReadSpecificColumnsFromRange<T>(sheetName, Dictionary<headerName, propertyName>, headerRow, headerRange)` — reads rows below the header row, maps Excel column headers to model properties by name.
- `WriteToRangeFromCollection<T>(sheetName, range, collection)` — writes a collection to a cell range.
- `SetFormula(sheetName, cell, formula)` — sets an Excel formula on a cell.
- `UpdateCell(sheetName, cell, value)` — updates a single named cell (e.g. `"B5"`).
- `Recalculate(sheetName, range)` — forces formula recalculation on a range.
- `ReadCellText(sheetName, cell)` — reads a cell value as string.
- `GeneratePdfFileFromWorkSheets(sheets[], ...)` — exports worksheets as PDFs.
- `Save()` / `Dispose()`.

For chart export, `Spire.Xls.Workbook` is used separately (not EPPlus):
- Load workbook → get sheet by name → get `Charts[0]` → `chart.RefreshChart()` → `chart.SaveToImage(path)`.

---

## Full User Workflow — Tab 3

### Step 1: Select Project
- A `<select>` dropdown bound to `selectedProject` string field, populated `OnInitializedAsync` from `AppSettings.Projects`.
- Changing selection calls `OnProjectChanged()` → loads sprint names.

### Step 2: Select Sprint
- A second `<select>` bound to `selectedSprint`, enabled only after project is chosen.
- Source: `SprintReportService.GetSprintNames(projectName)` — reads the `Data` sheet of `DataFileName`, returns non-empty Sprint values.
- Changing selection calls `OnSprintChanged()` → loads metrics.

### Step 3: Load Sprint Metrics into Data Table
- Sprint metrics displayed in a Blazor table (or MudBlazor `<MudDataGrid>`) with editable cells for: Committed, Delivered, Velocity, Code Review Comments (Internal/External), QA Defects, Escaped Defects, Backlog Health, Remarks.
- Source: `SprintReportService.GetSprintMetrics(projectName, sprintName)`:
  1. First tries `GetReportDataForProject(projectName)` — reads the `Data` sheet.
  2. If found, returns the matching row directly.
  3. If not found, reads from the per-team `MetricsSheetPath` Excel files (Dashboard sheet), aggregates across multiple team files (sum Committed/Delivered, sum defects), and computes:
     - `CommitmentIndex` = Committed / Delivered
     - `Velocity` = average of last 3 sprints from the Data sheet
     - `CodeQualityIndex` = sum of internal + external code review comments as `"{n}%"`
     - Defect counts extracted from `Remarks` field using pattern `"Label - {number}"`

### Step 4: Load Sprint Data (Click "Load Sprint Data" button)
This is the **core data processing step** — calls `SprintReportService.GenerateChart(projectName, sprintMetrics)`:

1. Validate: exactly 1 metrics record present.
2. Write metrics row to `Data` sheet:
   - Find the row index of the sprint in the Data sheet.
   - Write the row using `ExcelWrapper.WriteToRangeFromCollection("Data", "A{n}:M{n}", data)`.
   - Set formula on CommitmentIndex: `=IF((C{n}/B{n})>1,1,(C{n}/B{n}))`.
3. Update Scorecard:
   - Set cell `B5` = sprint name.
   - Recalculate range `B8:B13`.
   - Read score from cell `B3` (e.g. `"0.92"`).
4. Save the workbook.
5. Export chart image from `Scorecard` sheet using Spire.Xls → save to `wwwroot/downloads/{chartTitle}.png` (accessible via `<img src="/downloads/{chartTitle}.png">`).
6. Check if a PPT report already exists for this sprint; if yes, load and pre-fill text fields from slide 2 (shapes 6, 4, 7 for Summary, Highlights, Retrospective).
7. Bind to Blazor component state:
   - `chartImageUrl` = `/downloads/{chartTitle}.png`
   - `scoreText` = `"92%"`, `scoreCssClass` = `"score-green"` / `"score-yellow"` / `"score-white"`
   - `sprintSummaryLines`, `sprintHighlightLines`, `sprintRetrospectiveLines` (string arrays → joined with `\n` for textarea binding)

### Step 5: Edit Text Fields (optional)
Three `<textarea>` elements (or MudBlazor `<MudTextField Multiline="true">`):
- **Sprint Delivery Summary** — bullet-point lines
- **Sprint Highlights** — bullet-point lines
- **Sprint Retrospective** — structured text: "What went well\n...\nWhat didn't go well\n...\nImprovements\n..."

### Step 6: Generate Report (Click "Generate Report" button)
Validates all fields filled. Calls `SprintReportService.GeneratePresentation(ReportDataParameters)` which:
1. Loads the PPT template from `SprintMetricsReportTemplatePath`.
2. **Slide 1** (Title slide): Shape[0] gets 3 paragraphs:
   - "Delivery Quality Summary Report" (font 28, bold)
   - "{ProjectName} - " (font 18)
   - "{SprintNameWithDate} - " (font 12)
3. **Slide 2** (Content slide):
   - Shape[3]: Project name (plain text)
   - Shape[6]: Sprint Delivery Summary — bold header "Sprint Delivery Summary" (font 12), then each line as a bullet point (font 10, symbol bullet, indent -15)
   - Shape[4]: Highlights — bold header "Highlights:" then bullet points
   - Shape[7]: Retrospective — bold header "Retrospective:", then 3 sub-sections:
     - "What went well?" (underlined header, font 11) + bullet items
     - "What didn't go well?" + bullet items
     - "Improvements:" + bullet items
     - _(Category detection: line starts with "what went well", "what didnt go well", or "Improvements")_
     - _(Numbered list prefixes stripped: "1. item" → "item")_
   - Shape[8]: Sprint score text
   - Appends chart image at fixed position `(335, 26)` with size `390×180`, no border
4. Saves `.pptx` to project report folder: `GlobalPayments-{ProjectName}-DeliveryQualitySummaryReport-{SprintName}.pptx`
5. Copies `.pdf` to `wwwroot/downloads/` for browser download
6. Returns `(true, pdfRelativeUrl)` — binds `pdfDownloadUrl` to state

### Step 7: Download / View PDF
- A `<a href="@pdfDownloadUrl" target="_blank">View PDF</a>` link appears after report generation.
- The PDF is served from `wwwroot/downloads/` as a static file.
- Alternatively, use Blazor's `IJSRuntime` to trigger a file download.

### Step 8: View Email Content
- A collapsible panel (or modal dialog using MudBlazor `<MudDialog>`) displays the PMO email HTML content, generated by `GetEmailContent`:
  - Takes `PMOEmailContent` template from config.
  - Replaces `##` with the report filename.
  - Replaces `#Link#` with `{ProjectFolderOneDriveLink}{reportFilename}`.
- Rendered with `@((MarkupString)emailContent)`.

---

## Blazor Page Structure (`SprintReport.razor`)

```razor
@page "/sprint-report"
@inject SprintReportService Service
@inject IOptions<AppSettings> AppSettings

<!-- Step 1: Project & Sprint Selection -->
<section>
  <select @onchange="OnProjectChanged">...</select>
  <select @onchange="OnSprintChanged" disabled="@(sprintNames == null)">...</select>
</section>

<!-- Step 2: Metrics Data Grid (editable) -->
@if (sprintMetrics != null)
{
  <table><!-- or MudDataGrid --></table>
  <button @onclick="LoadSprintData">Load Sprint Data</button>
}

<!-- Step 3: Chart + Score + Text fields (shown after Load) -->
@if (chartImageUrl != null)
{
  <div class="layout-grid">
    <img src="@chartImageUrl" />
    <span class="@scoreCssClass">@scoreText</span>
    <textarea @bind="sprintSummary" />
    <textarea @bind="sprintHighlights" />
    <textarea @bind="sprintRetrospective" />
  </div>
  <button @onclick="GenerateReport">Generate Report</button>
}

<!-- Step 4: Download + Email (shown after report generation) -->
@if (pdfDownloadUrl != null)
{
  <a href="@pdfDownloadUrl" target="_blank">View / Download PDF</a>
  <button @onclick="ShowEmailContent">Email Content</button>
}
```

CSS classes for score: `.score-green { background: green; color: white; }`, `.score-yellow { background: yellow; }`, `.score-white { background: white; }`.

---

## UI Layout (Blazor Page Controls)

| Element | Blazor / HTML | Purpose |
|---|---|---|
| Project dropdown | `<select>` / `MudSelect` | Project selection |
| Sprint dropdown | `<select>` / `MudSelect` | Sprint selection |
| Metrics table | `<table>` / `MudDataGrid` | Sprint metrics display/edit |
| Load Sprint Data button | `<button>` / `MudButton` | Triggers chart + score generation |
| Chart image | `<img src="@chartImageUrl">` | Displays exported chart PNG |
| Score badge | `<span class="@scoreCssClass">@scoreText</span>` | Color-coded score display |
| Sprint Summary textarea | `<textarea @bind>` / `MudTextField Multiline` | Sprint Summary input |
| Sprint Highlights textarea | `<textarea @bind>` / `MudTextField Multiline` | Highlights input |
| Sprint Retrospective textarea | `<textarea @bind>` / `MudTextField Multiline` | Retrospective input |
| Generate Report button | `<button>` / `MudButton` | Generates PPT + PDF |
| View PDF link | `<a href target="_blank">` | Opens/downloads PDF |
| Email Content button | `<button>` / `MudButton` | Shows PMO email in dialog |

---

## Supporting Extension Methods (`SprintReport.Data/Extensions`)

- `StringExtensions.ToEmptyStringIfNull()` — returns `""` if null
- `StringExtensions.ToLong()`, `.ToInt()`, `.ToDouble()` — safe object-to-numeric conversions
- `StringExtensions.RemoveEmptyLines()` — filters empty strings from a string array
- `FileExtensions.GetRecentlyModifiedSimilarFile()` — given a file path, returns the most recently modified file in the same directory with a matching name pattern

---

## Todos

1. **Model layer** — Define `SprintMetrics`, `PPTReportData`, `ReportDataParameters`, `AppSettings`, `Project`, `DashboardData`, `EmailParameters` POCOs in `SprintReport.Model`
2. **Data layer / Config** — `AppConfiguration` binding; use `IOptions<AppSettings>` in `Program.cs`; `PopulateDirectories()` helper; Serilog file logger setup via `UseSerilog`
3. **Data layer / Extensions** — `StringExtensions`, `FileExtensions`, `ObjectExtensions` in `SprintReport.Data/Extensions`
4. **Data layer / ExcelWrapper** — `IWrapper` interface + `ExcelWrapper` implementation (EPPlus for read/write/formula/recalculate; Spire.Xls for chart-to-PNG export to `wwwroot/downloads/`)
5. **Data layer / DataProcessor** — Implement `IDataProcessor`:
   - `GetSprintNames` — reads Data sheet sprint column
   - `GetSprintMetrics` — reads from Data sheet or falls back to metrics sheets + aggregation + defect parsing
   - `GenerateChart` — writes to Excel, recalculates scorecard, exports chart PNG, reads back existing PPT content
   - `GeneratePresentation` — builds PPT from template (3 slides, shaped text, bullet formatting, retrospective sections, chart image embed), exports PDF to `wwwroot/downloads/`
   - `GetEmailContent` — builds PMO email HTML from config template
6. **Service layer** — `SprintReportService` registered as `AddScoped` in DI; thin delegation to `IDataProcessor`; injected into Blazor pages
7. **Blazor page** — `SprintReport.razor` at route `/sprint-report` with: project/sprint selects, editable metrics table, Load button, chart image display, color-coded score badge, 3 textareas, Generate Report button, PDF download link, Email Content modal
8. **Blazor layout + nav** — `MainLayout.razor`, `NavMenu.razor` with link to Sprint Report page; Bootstrap or MudBlazor styling
9. **Configuration + assets** — `appsettings.json` with `AppSettings` section; copy PPT template to `Templates/` folder in web project; configure static file serving for `wwwroot/downloads/`
