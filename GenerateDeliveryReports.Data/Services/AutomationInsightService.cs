using System.Text.Json;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace GenerateDeliveryReports.Data.Services;

public interface IAutomationInsightService
{
    Task<AutomationInsightResult> BuildAsync(AutomationWorkbookReport report, CancellationToken cancellationToken = default);
}

public sealed class AutomationInsightService(IChatClient chatClient, IOptions<AutomationReportingOptions> options, ILogger<AutomationInsightService> logger) : IAutomationInsightService
{
    public async Task<AutomationInsightResult> BuildAsync(AutomationWorkbookReport report, CancellationToken cancellationToken = default)
    {
        var facts = BuildFacts(report);
        var highlights = facts.Take(options.Value.Pptx.MaxFooterHighlights)
            .Select(fact => new AutomationHighlight(fact.Project, fact.Category, fact.Text, [fact.Id]))
            .ToList();
        var comments = report.Projects.Select(project => new CommentAssessment(project.Current.Project, project.Comment,
            CommentClassification.ReasoningOrContext, [], null, false)).ToList();

        if (!options.Value.Llm.Enabled || !comments.Any(comment => !string.IsNullOrWhiteSpace(comment.OriginalComment)))
            return new AutomationInsightResult(facts, highlights, comments, false, null);

        try
        {
            var prompt = BuildCommentPrompt(report, facts);
            var response = await chatClient.GetResponseAsync(
                [new ChatMessage(ChatRole.System, "Return only JSON. Classify team comments. Use only supplied fact IDs. Do not infer causation or invent values."), new ChatMessage(ChatRole.User, prompt)],
                new ChatOptions { Temperature = 0, MaxOutputTokens = options.Value.Llm.Ollama.MaxTokens }, cancellationToken);
            var modelComments = ParseAndValidateComments(response.Text, report, facts);
            return new AutomationInsightResult(facts, highlights, modelComments, true, null);
        }
        catch (Exception ex) when (options.Value.Llm.FallbackToDeterministicInsights)
        {
            logger.LogWarning(ex, "Automation report LLM failed; deterministic insights will be used.");
            return new AutomationInsightResult(facts, highlights, comments, false, ex.Message);
        }
    }

    private static List<VerifiedFact> BuildFacts(AutomationWorkbookReport report)
    {
        var conflicts = report.Validations.Where(validation => validation.Severity == ValidationSeverity.Conflict)
            .Select(validation => validation.Project).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var facts = new List<VerifiedFact>();
        foreach (var project in report.Projects.Where(project => !conflicts.Contains(project.Current.Project)))
        {
            var comparison = project.Comparison;
            if (comparison?.CoveragePercentagePointChange is decimal coverageDelta && coverageDelta != 0 && project.Current.CalculatedCoverage is decimal coverage)
                facts.Add(new VerifiedFact($"{project.Current.Project}:coverage", project.Current.Project, "Coverage", $"{project.Current.Project} coverage {(coverageDelta > 0 ? "increased" : "decreased")} {Math.Abs(coverageDelta * 100):0.0} percentage points to {coverage * 100:0.0}% compared with the previous week."));
            if (comparison?.AutomatedTestChange is decimal automatedDelta && automatedDelta != 0 && project.Current.Automated is decimal automated)
                facts.Add(new VerifiedFact($"{project.Current.Project}:automated", project.Current.Project, "Automation delivery", $"{project.Current.Project} automated test inventory {(automatedDelta > 0 ? "increased" : "decreased")} by {Math.Abs(automatedDelta):0} to {automated:0}."));
            if (comparison?.PassRatePercentagePointChange is decimal passDelta && passDelta != 0 && project.Current.CalculatedPassRate is decimal passRate)
                facts.Add(new VerifiedFact($"{project.Current.Project}:run-rate", project.Current.Project, "Run performance", $"{project.Current.Project} pass rate {(passDelta > 0 ? "increased" : "decreased")} {Math.Abs(passDelta * 100):0.0} percentage points to {passRate * 100:0.0}% compared with the previous week."));
        }
        return facts;
    }

    private static string BuildCommentPrompt(AutomationWorkbookReport report, IReadOnlyCollection<VerifiedFact> facts) => JsonSerializer.Serialize(new
    {
        comments = report.Projects.Select(project => new { project = project.Current.Project, comment = project.Comment }),
        facts = facts.Select(fact => new { fact.Id, fact.Project, fact.Text }),
        output = "Return [{ project, classification, factIds, suggestedWording }]. classification must be DataSubstantiation, ReasoningOrContext, UnsupportedNumericClaim, or UnverifiableContext."
    });

    private static IReadOnlyList<CommentAssessment> ParseAndValidateComments(string? raw, AutomationWorkbookReport report, IReadOnlyCollection<VerifiedFact> facts)
    {
        var allowedIds = facts.Select(fact => fact.Id).ToHashSet(StringComparer.Ordinal);
        var byProject = report.Projects.ToDictionary(project => project.Current.Project, StringComparer.OrdinalIgnoreCase);
        var candidates = JsonSerializer.Deserialize<List<ModelComment>>(raw ?? "[]", new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];
        return candidates.Where(candidate => !string.IsNullOrWhiteSpace(candidate.Project) && byProject.ContainsKey(candidate.Project))
            .Select(candidate => new CommentAssessment(candidate.Project!, byProject[candidate.Project!].Comment,
                Enum.TryParse<CommentClassification>(candidate.Classification, true, out var classification) ? classification : CommentClassification.UnverifiableContext,
                (candidate.FactIds ?? []).Where(allowedIds.Contains).ToList(), Truncate(candidate.SuggestedWording, 185), true)).ToList();
    }

    private static string? Truncate(string? text, int maxLength) => string.IsNullOrWhiteSpace(text) ? null : text.Length <= maxLength ? text : text[..(maxLength - 1)] + "…";
    private sealed class ModelComment { public string? Project { get; set; } public string? Classification { get; set; } public List<string>? FactIds { get; set; } public string? SuggestedWording { get; set; } }
}
