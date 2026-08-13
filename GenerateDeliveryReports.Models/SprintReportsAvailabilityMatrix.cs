namespace GenerateDeliveryReports.Models;

/// <summary>
/// Root object in the SprintReportsAvailability.json file.
/// Shows all products and their report generation status by sprint.
/// </summary>
public class SprintReportsAvailabilityMatrix
{
    /// <summary>ISO date when generate-sprint-reports-availability.ps1 was last run.</summary>
    public string? Generated { get; set; }

    /// <summary>The cutoff date used when generating (e.g. "2026-04-01").</summary>
    public string? FromDate { get; set; }

    /// <summary>List of products with their sprint report entries.</summary>
    public List<ProductReportMatrix> Matrix { get; set; } = [];
}

/// <summary>
/// Product entry containing all sprints and their report status.
/// </summary>
public class ProductReportMatrix
{
    /// <summary>Product/Application name (e.g., "Hosted Payments").</summary>
    public string? Application { get; set; }

    /// <summary>List of sprint report entries for this product.</summary>
    public List<SprintReportEntry> Entries { get; set; } = [];
}

/// <summary>
/// Individual sprint report entry with lag day calculation and color status.
/// </summary>
public class SprintReportEntry
{
    /// <summary>Sprint name without date range (e.g., "Sprint 12").</summary>
    public string? SprintName { get; set; }

    /// <summary>ISO format sprint start date (yyyy-MM-dd).</summary>
    public string? SprintStartDate { get; set; }

    /// <summary>ISO format sprint end date (yyyy-MM-dd).</summary>
    public string? SprintEndDate { get; set; }

    /// <summary>ISO format date when the report was created/written (yyyy-MM-dd). Null if missing.</summary>
    public string? ReportCreatedDate { get; set; }

    /// <summary>
    /// Days between sprint end date and report creation.
    /// Null if report is missing.
    /// Positive values indicate delay (e.g., +1, +3).
    /// </summary>
    public int? LagDays { get; set; }

    /// <summary>
    /// Status/Color indicator:
    /// - "Green" = 0 days lag (same day)
    /// - "Yellow" = 1-2 days lag
    /// - "Red" = 3+ days lag
    /// - "Missing" = No report found
    /// </summary>
    public string? Status { get; set; }

    /// <summary>Non-null only when sprint name could not be parsed. Describes the parsing error.</summary>
    public string? ParseError { get; set; }

    /// <summary>Expected filename that was searched for (helps debug why reports are marked as missing).</summary>
    public string? ExpectedFileName { get; set; }

    /// <summary>Actual filename that was matched (may differ from expected if fuzzy matching was used).</summary>
    public string? ActualFileName { get; set; }
}
