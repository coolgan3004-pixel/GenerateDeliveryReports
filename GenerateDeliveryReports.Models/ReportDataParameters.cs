namespace GenerateDeliveryReports.Models;

public class ReportDataParameters
{
    public string ProjectName { get; set; } = string.Empty;
    public string SprintNameWithDate { get; set; } = string.Empty;
    public string SprintName => SprintNameWithDate.Contains('(')
        ? SprintNameWithDate[..SprintNameWithDate.IndexOf('(')].Trim()
        : SprintNameWithDate.Trim();
    public string ImagePath { get; set; } = string.Empty;
    public string[]? SprintHighlights { get; set; }
    public string[]? SprintSummary { get; set; }
    public string[]? SprintRetrospective { get; set; }
    public string SprintScore { get; set; } = string.Empty;
}
