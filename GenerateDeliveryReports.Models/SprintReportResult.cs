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
}
