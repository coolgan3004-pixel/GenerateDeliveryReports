using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Models;

public class RunHistory
{
    public DateTime RunTime { get; set; }
    public List<SprintReportResult> Results { get; set; } = new();

    public int CompletedCount => Results.Count(r => r.Outcome == SprintReportOutcome.Completed);    
    public int ErrorCount => Results.Count(r => r.Outcome == SprintReportOutcome.Errored);    
    public int SkippedCount => Results.Count(r => r.Outcome == SprintReportOutcome.Missing);    

    // TODO: Add properties for counts
    // Tip: Use LINQ to count based on result.Outcome
}