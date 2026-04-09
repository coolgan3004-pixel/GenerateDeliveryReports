using GenerateDeliveryReports.Data.Interface;
using SendGrid;
using SendGrid.Helpers.Mail;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Options;

namespace GenerateDeliveryReports.Worker;

public class ReportWorker : BackgroundService
{
    private readonly IDataProcessor _dataProcessor;
    private readonly AppSettings _appSettings;
    private readonly ILogger<ReportWorker> _logger;

    public ReportWorker(IDataProcessor dataProcessor, IOptions<AppSettings> appSettings, ILogger<ReportWorker> logger)
    {
        _dataProcessor = dataProcessor;
        _appSettings = appSettings.Value;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        if (_appSettings.WorkerIntervalMinutes <= 0)
        {
            _logger.LogError("WorkerIntervalMinutes is not configured or is invalid. Worker will not run.");
            return;
        }

        _logger.LogInformation("ReportWorker started. Interval: {Interval} minutes.", _appSettings.WorkerIntervalMinutes);

        while (!stoppingToken.IsCancellationRequested)
        {
            var cycleTime = DateTimeOffset.Now;
            _logger.LogInformation("ReportWorker cycle starting at {Time}", cycleTime);

            var results = await ProcessAllProjectsAsync(stoppingToken);
            await SaveCycleSummaryAsync(results, cycleTime);

            _logger.LogInformation("ReportWorker cycle complete. Next run in {Interval} minutes.", _appSettings.WorkerIntervalMinutes);
            await Task.Delay(TimeSpan.FromMinutes(_appSettings.WorkerIntervalMinutes), stoppingToken);
        }
    }

    private async Task<List<SprintReportResult>> ProcessAllProjectsAsync(CancellationToken stoppingToken)
    {
        var results = new List<SprintReportResult>();

        IEnumerable<string> projects;
        try
        {
            projects = _dataProcessor.GetProjectNames();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to get project names.");
            return results;
        }

        foreach (var projectName in projects)
        {
            if (stoppingToken.IsCancellationRequested) break;
            results.AddRange(await ProcessProjectAsync(projectName, stoppingToken));
        }

        return results;
    }

    private async Task<IEnumerable<SprintReportResult>> ProcessProjectAsync(string projectName, CancellationToken stoppingToken)
    {
        _logger.LogInformation("Processing project: {Project}", projectName);

        IEnumerable<SprintInfo> sprints;
        try
        {
            sprints = _dataProcessor.GetSprintNamesWithDate(projectName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to get sprint names for project {Project}.", projectName);
            return [];
        }

        var results = new List<SprintReportResult>();
        foreach (var sprint in sprints)
        {
            if (stoppingToken.IsCancellationRequested) break;
            var result = await ProcessSprintAsync(sprint);
            if (result != null) results.Add(result);
        }

        return results;
    }

    private async Task<SprintReportResult?> ProcessSprintAsync(SprintInfo sprint)
    {
        // Skip if sprint end date has not passed yet
        if (sprint.SprintEndDate == null)
        {
            _logger.LogWarning("[{Project}] [{Sprint}] Skipping — could not determine sprint end date.", sprint.ProjectName, sprint.SprintName);
            return null;
        }

        if (sprint.SprintEndDate.Value > DateTime.Today)
        {
            _logger.LogInformation("[{Project}] [{Sprint}] Skipping — sprint ends {EndDate}, not yet due.", sprint.ProjectName, sprint.SprintName, sprint.SprintEndDate.Value.ToShortDateString());
            return null;
        }

        if (sprint.SprintEndDate.Value < new DateTime(2026, 4, 1))
        {
            _logger.LogInformation("[{Project}] [{Sprint}] Skipping — sprint ended before Apr 1 2026.", sprint.ProjectName, sprint.SprintName);
            return null;
        }

        // Skip if report already exists
        if (File.Exists(sprint.OutputPPTPath))
        {
            _logger.LogInformation("[{Project}] [{Sprint}] Skipping — report already exists at {Path}.", sprint.ProjectName, sprint.SprintName, sprint.OutputPPTPath);
            return null;
        }

        // Report is missing for a completed sprint
        _logger.LogWarning("[{Project}] [{Sprint}] Missing — no report found for completed sprint.", sprint.ProjectName, sprint.SprintName);

        var metrics = FetchSprintMetrics(sprint);
        if (metrics == null)
            return new SprintReportResult { ProjectName = sprint.ProjectName, SprintName = sprint.SprintName, SprintEndDate = sprint.SprintEndDate, Outcome = SprintReportOutcome.Errored, Detail = $"Metrics unavailable or narrative fields incomplete. Expected: {sprint.OutputPPTPath}" };
        return await CreateReportAsync(sprint, metrics);
    }

    private SprintMetrics? FetchSprintMetrics(SprintInfo sprint)
    {
        SprintMetrics? metrics;
        try
        {
            metrics = _dataProcessor.GetSprintMetrics(sprint.ProjectName, sprint.SprintName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[{Project}] [{Sprint}] Failed to get sprint metrics.", sprint.ProjectName, sprint.SprintName);
            return null;
        }

        if (metrics == null)
        {
            _logger.LogWarning("[{Project}] [{Sprint}] Skipping — no metrics returned.", sprint.ProjectName, sprint.SprintName);
            return null;
        }

        // Validate all 3 narrative fields are non-empty
        var summaryMissing = metrics.SprintSummary == null || !metrics.SprintSummary.Any(s => !string.IsNullOrWhiteSpace(s));
        var highlightsMissing = metrics.SprintHighlights == null || !metrics.SprintHighlights.Any(s => !string.IsNullOrWhiteSpace(s));
        var retrospectiveMissing = metrics.SprintRetrospective == null || !metrics.SprintRetrospective.Any(s => !string.IsNullOrWhiteSpace(s));

        if (summaryMissing || highlightsMissing || retrospectiveMissing)
        {
            _logger.LogWarning(
                "[{Project}] [{Sprint}] Skipping report generation — one or more narrative fields are empty. " +
                "Summary: {S}, Highlights: {H}, Retrospective: {R}",
                sprint.ProjectName, sprint.SprintName,
                summaryMissing ? "MISSING" : "OK",
                highlightsMissing ? "MISSING" : "OK",
                retrospectiveMissing ? "MISSING" : "OK");
            return null;
        }

        return metrics;
    }

    private async Task<SprintReportResult> CreateReportAsync(SprintInfo sprint, SprintMetrics metrics)
    {
        _logger.LogInformation("[{Project}] [{Sprint}] Generating report...", sprint.ProjectName, sprint.SprintName);
        var reportParams = new ReportDataParameters
        {
            ProjectName = sprint.ProjectName,
            SprintNameWithDate = sprint.SprintName,
            ImagePath = metrics.ImagePath,
            SprintScore = metrics.Score,
            SprintSummary = metrics.SprintSummary,
            SprintHighlights = metrics.SprintHighlights,
            SprintRetrospective = metrics.SprintRetrospective
        };

        try
        {
            var (success, outputPath) = _dataProcessor.GeneratePresentation(reportParams, generatePdf: false);
            if (success)
            {
                _logger.LogInformation("[{Project}] [{Sprint}] Report generated: {Path}", sprint.ProjectName, sprint.SprintName, outputPath);
                return new SprintReportResult { ProjectName = sprint.ProjectName, SprintName = sprint.SprintName, SprintEndDate = sprint.SprintEndDate, Outcome = SprintReportOutcome.Completed, Detail = outputPath };
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[{Project}] [{Sprint}] Failed to generate report.", sprint.ProjectName, sprint.SprintName);
            return new SprintReportResult { ProjectName = sprint.ProjectName, SprintName = sprint.SprintName, SprintEndDate = sprint.SprintEndDate, Outcome = SprintReportOutcome.Errored, Detail = ex.Message };
        }

        return new SprintReportResult { ProjectName = sprint.ProjectName, SprintName = sprint.SprintName, SprintEndDate = sprint.SprintEndDate, Outcome = SprintReportOutcome.Errored, Detail = "GeneratePresentation returned false with no exception." };
    }

    private async Task SaveCycleSummaryAsync(List<SprintReportResult> results, DateTimeOffset cycleTime)
    {
        var html = ReportEmailBuilder.BuildHtml(results, cycleTime);

        try
        {
            var filePath = string.IsNullOrWhiteSpace(_appSettings.WorkerSummaryFilePath)
                ? Path.Combine(Directory.GetCurrentDirectory(), "LogFiles", "worker-summary.html")
                : _appSettings.WorkerSummaryFilePath;

            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);
            await File.WriteAllTextAsync(filePath, html);
            _logger.LogInformation("Cycle summary saved: {Path}", filePath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to save cycle summary.");
        }

        await SendCycleSummaryEmailAsync(html, cycleTime);
    }

    private async Task SendCycleSummaryEmailAsync(string html, DateTimeOffset cycleTime)
    {
        var email = _appSettings.EmailSettings;

        var apiKey = string.IsNullOrWhiteSpace(email.Password)
            ? Environment.GetEnvironmentVariable("SENDGRID_API_KEY")
            : email.Password;
        _logger.LogInformation("SendGrid — key length: {Len}, from: {From}, recipients: {Count}",
            apiKey?.Length ?? 0, email.FromEmailAddress, email.Users.Count);

        if (string.IsNullOrWhiteSpace(email.SMTPServer) ||
            string.IsNullOrWhiteSpace(apiKey) ||
            string.IsNullOrWhiteSpace(email.FromEmailAddress) ||
            email.Users.Count == 0)
        {
            _logger.LogWarning("Email not sent — SMTP settings or recipient list is incomplete.");
            return;
        }

        try
        {
            var client = new SendGridClient(apiKey);
            var message = new SendGridMessage
            {
                From = new EmailAddress(email.FromEmailAddress),
                Subject = $"Delivery Report Cycle Summary — {cycleTime:yyyy-MM-dd HH:mm}",
                HtmlContent = html
            };

            foreach (var user in email.Users)
                message.AddTo(new EmailAddress(user.Email, user.Name));

            var response = await client.SendEmailAsync(message);

            if ((int)response.StatusCode >= 200 && (int)response.StatusCode < 300)
                _logger.LogInformation("Cycle summary email sent to {Count} recipient(s).", email.Users.Count);
            else
            {
                var body = await response.Body.ReadAsStringAsync();
                _logger.LogError("SendGrid returned {Status}: {Body}", response.StatusCode, body);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to send cycle summary email.");
        }
    }
}
