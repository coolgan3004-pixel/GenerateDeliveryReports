namespace GenerateDeliveryReports.Models;

public enum SprintReportOutcome
{
    Missing,
    Completed,
    Errored
}

public class SprintReportResult
{
    public string ProjectName { get; set; } = string.Empty;
    public string SprintName { get; set; } = string.Empty;
    public SprintReportOutcome Outcome { get; set; }
    public DateTime? SprintEndDate { get; set; }
    public string? Detail { get; set; }

    // Populated for Errored results where metrics were fetched
    public string? SprintMetricsSprintName { get; set; }
    public string[]? SprintSummary { get; set; }
    public string[]? SprintHighlights { get; set; }
    public string[]? SprintRetrospective { get; set; }
}
