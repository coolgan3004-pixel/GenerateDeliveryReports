namespace GenerateDeliveryReports.Models;

public sealed class AutomationReportingOptions
{
    public LlmOptions Llm { get; set; } = new();
    public PptxOptions Pptx { get; set; } = new();
}

public sealed class LlmOptions
{
    public bool Enabled { get; set; }
    public string Provider { get; set; } = "Ollama";
    public bool FallbackToDeterministicInsights { get; set; } = true;
    public int RequestTimeoutSeconds { get; set; } = 45;
    public int MaxConcurrentRequests { get; set; } = 1;
    public OllamaOptions Ollama { get; set; } = new();
}

public sealed class OllamaOptions
{
    public string BaseUrl { get; set; } = "http://localhost:11434";
    public string Model { get; set; } = "qwen2.5:3b-instruct";
    public double Temperature { get; set; } = 0.1;
    public int MaxTokens { get; set; } = 350;
}

public sealed class PptxOptions { public int MaxFooterHighlights { get; set; } = 3; }

public sealed record AutomationReportSelection(string CurrentSheetName, string PreviousSheetName, IReadOnlyDictionary<string, MetricResolution>? Resolutions = null);
public enum MetricResolution { UseCalculated, UseEntered }
public enum ValidationSeverity { Warning, Conflict, Error }
public enum CommentClassification { DataSubstantiation, ReasoningOrContext, UnsupportedNumericClaim, UnverifiableContext }

public sealed class AutomationWorkbookReport
{
    public required string Client { get; init; }
    public required string Team { get; init; }
    public required string CurrentSheetName { get; init; }
    public required DateOnly CurrentWeek { get; init; }
    public string? PreviousSheetName { get; init; }
    public DateOnly? PreviousWeek { get; init; }
    public List<AutomationProjectReport> Projects { get; init; } = [];
    public List<CalculatedMetricValidation> Validations { get; init; } = [];
    public List<string> Warnings { get; init; } = [];
}

public sealed class AutomationProjectReport
{
    public required AutomationProjectMetrics Current { get; init; }
    public AutomationProjectMetrics? Previous { get; init; }
    public AutomationWeeklyComparison? Comparison { get; init; }
    public string Comment => Current.Comment ?? string.Empty;
}

public sealed class AutomationProjectMetrics
{
    public int SourceRow { get; init; }
    public required string Project { get; init; }
    public string? Sprint { get; init; }
    public string? Status { get; init; }
    public decimal? ManualTests { get; init; }
    public decimal? CanBeAutomated { get; init; }
    public decimal? Automated { get; init; }
    public decimal? AddedThisSprint { get; init; }
    public decimal? TargetMinimum { get; init; }
    public decimal? TargetMaximum { get; init; }
    public decimal? Passed { get; init; }
    public decimal? TimedOut { get; init; }
    public decimal? Failed { get; init; }
    public decimal? LowDefects { get; init; }
    public decimal? MediumDefects { get; init; }
    public decimal? HighDefects { get; init; }
    public string? Comment { get; init; }
    public decimal? EnteredCoverage { get; init; }
    public decimal? EnteredPassRate { get; init; }
    public decimal? CalculatedCoverage { get; init; }
    public decimal? CalculatedPassRate { get; init; }
}

public sealed record AutomationWeeklyComparison(decimal? CoveragePercentagePointChange, decimal? PassRatePercentagePointChange, decimal? AutomatedTestChange, decimal? AddedThisSprintChange);
public sealed record CalculatedMetricValidation(string Project, string Cell, string Metric, decimal? EnteredValue, decimal? CalculatedValue, string SourceType, ValidationSeverity Severity, string Message);
public sealed record VerifiedFact(string Id, string Project, string Category, string Text);
public sealed record AutomationHighlight(string Project, string Category, string Text, IReadOnlyList<string> FactIds, bool IsApproved = false);
public sealed record CommentAssessment(string Project, string OriginalComment, CommentClassification Classification, IReadOnlyList<string> FactIds, string? SuggestedWording, bool IsModelGenerated);
public sealed record AutomationInsightResult(IReadOnlyList<VerifiedFact> Facts, IReadOnlyList<AutomationHighlight> Highlights, IReadOnlyList<CommentAssessment> Comments, bool UsedLlm, string? ProviderError);
