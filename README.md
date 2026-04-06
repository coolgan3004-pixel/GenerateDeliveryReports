# GenerateDeliveryReports

A Blazor Server application for generating Sprint Delivery Quality Summary Reports. It reads sprint metrics from Excel files stored on OneDrive, generates PPTX and PDF reports from a template, and provides email content for distribution.

## Tech Stack

- **.NET 10** / **Blazor Server** (Interactive Server rendering)
- **Spire.XLS** — Excel file reading and chart image export
- **Spire.Presentation** — PowerPoint report generation (PPTX & PDF)
- **EPPlus** — Additional Excel processing
- **Serilog** — Structured logging with daily rolling file output

## Solution Structure

```
GenerateDeligeryReports.sln
├── GenerateDeliveryReports/              # Blazor Server web app
│   ├── Components/
│   │   ├── Layout/                       # MainLayout, NavMenu, ReconnectModal
│   │   └── Pages/
│   │       ├── Home.razor
│   │       └── GenerateReport.razor      # Main report generation page
│   ├── wwwroot/                          # Static assets
│   │   └── downloads/                    # Generated chart images & PDFs (git-ignored)
│   ├── LogFiles/                         # Serilog rolling log files
│   ├── Program.cs                        # App startup & DI configuration
│   └── appsettings.json                  # App configuration
│
├── GenerateDeliveryReports.Data/         # Data access & business logic
│   ├── Concrete/
│   │   └── DataProcessor.cs              # Core report processing logic
│   ├── Interface/
│   │   └── IDataProcessor.cs             # Data processor contract
│   ├── Services/
│   │   └── SprintReportService.cs        # Facade service for UI consumption
│   ├── Extensions/                       # Helper/extension methods
│   ├── Templates/                        # PPTX report template
│   └── Settings/                         # Configuration helpers
│
└── GenerateDeliveryReports.Models/       # Shared models
    ├── AppSettings / ProjectsConfig.cs   # Configuration models
    ├── SprintMetrics.cs                  # Sprint data model
    ├── ReportDataParameters.cs           # Report generation parameters
    ├── ErrorCodes.cs                     # Centralized error codes & messages
    └── ...
```

## Report Generation Workflow

1. **Select Project** — Loads available sprint names from the project's Excel data file.
2. **Select Sprint** — Fetches sprint metrics (committed, delivered, velocity, defects, etc.) and exports the scorecard chart image. Pre-fills narrative fields if a previous report exists.
3. **Review & Edit** — Displays metrics in an editable table with the chart image, score badge, and text fields for Summary, Highlights, and Retrospective.
4. **Generate Report** — Creates a PPTX from the template, populates slides with data and chart image, exports to PDF, and provides a download link plus email content.

## Configuration

Key settings in `appsettings.json` under `AppSettings`:

| Setting | Description |
|---|---|
| `OneDriveLocation` | Local OneDrive sync root path |
| `ReportAndDataFolder` | Subfolder containing project data Excel files |
| `MetricsFolder` | Subfolder containing team sprint metrics sheets |
| `SprintMetricsReportTemplatePath` | Full path to the PPTX report template |
| `PMOEmailContent` | Email body template with `##` and `#Link#` placeholders |
| `Projects` | Array of project configurations (name, metrics paths, data file, OneDrive link) |

## Getting Started

### Prerequisites

- .NET 10 SDK
- OneDrive folder synced locally with required Excel data files
- PPTX template file at the configured path

### Run

```bash
dotnet run --project GenerateDeliveryReports/GenerateDeliveryReports.csproj
```

The app launches at `https://localhost:5001` (or the port configured in `launchSettings.json`).

## Error Handling

The application uses centralized error codes defined in `ErrorCodes.cs`:

| Code | Description |
|---|---|
| `ERR_100` | Data file not found |
| `ERR_101` | No recently modified file found |
| `ERR_103` | PPT template file not found |
| `ERR_200` | No sprint data found |
| `ERR_203` | Chart not found in scorecard sheet |
| `ERR_300` | Invalid template (too many slides) |
| `ERR_301` | Failed to generate/save presentation |
| `ERR_303` | Chart image file not found |
| `ERR_400` | Project not found in configuration |

All errors surface in the UI as alert messages.
