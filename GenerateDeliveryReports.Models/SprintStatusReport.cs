namespace GenerateDeliveryReports.Models;

/// <summary>
/// Root object in the sprint-report-status.json file.
/// </summary>
public class SprintStatusReport
{
    /// <summary>ISO date when generate-sprint-status.ps1 was last run.</summary>
    public string? Generated { get; set; }

    /// <summary>The cutoff date used when generating (e.g. "2026-04-01").</summary>
    public string? FromDate { get; set; }

    public List<SprintStatusEntry> Entries { get; set; } = [];
}
