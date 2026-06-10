namespace GenerateDeliveryReports.Models;

/// <summary>
/// One row in the sprint-report-status.json file produced by generate-sprint-status.ps1.
/// </summary>
public class SprintStatusEntry
{
    /// <summary>"April 2026" — the month the sprint ended (null for parse errors).</summary>
    public string? Month { get; set; }

    public string? Application { get; set; }
    public string? SprintName { get; set; }
    public string? SprintStartDate { get; set; }
    public string? SprintEndDate { get; set; }

    /// <summary>Available | Missing | Upcoming | ParseError</summary>
    public string? ReportStatus { get; set; }

    public string? ReportCreatedDate { get; set; }

    /// <summary>Days between sprint end date and the report's last-write date (null when no report).</summary>
    public int? DaysDiff { get; set; }

    /// <summary>Non-null only when ReportStatus == "ParseError".</summary>
    public string? ParseError { get; set; }
}
