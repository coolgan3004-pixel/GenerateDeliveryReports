namespace GenerateDeliveryReports.Models;

public class PPTReportData
{
    public string ImagePath { get; set; } = string.Empty;
    public string Score { get; set; } = string.Empty;
    public string[]? SprintSummary { get; set; }
    public string[]? SprintHighlights { get; set; }
    public string[]? SprintRetrospective { get; set; }
}
