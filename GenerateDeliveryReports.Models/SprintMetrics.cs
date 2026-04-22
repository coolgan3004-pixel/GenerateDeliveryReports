namespace GenerateDeliveryReports.Models;

public class SprintMetrics
{
    public string Sprint { get; set; } = string.Empty;
    public object? Committed { get; set; }
    public object? Delivered { get; set; }
    public object? CommitmentIndex { get; set; }
    public object? Velocity { get; set; }
    public object? CodeQualityIndex { get; set; }
    public object? CodeReviewCommentsInternal { get; set; }
    public object? CodeReviewCommentsExternal { get; set; }
    public object? QADefects { get; set; }
    public object? EscapedDefects { get; set; }
    public object? BacklogHealth { get; set; }
    public string LastSprint { get; set; } = string.Empty;
    public string SprintNameFromMetricsSheet { get; set; } = string.Empty;
    public object? Remarks { get; set; }

     public string ImagePath { get; set; } = string.Empty;
    public string OutputPPTPath { get; set; } = string.Empty;
    public string Score { get; set; } = string.Empty;
    public string[]? SprintSummary { get; set; }
    public string SprintMetricsSprintName { get; set; } = string.Empty;
    public string[]? SprintHighlights { get; set; }
    public string[]? SprintRetrospective { get; set; }
}
