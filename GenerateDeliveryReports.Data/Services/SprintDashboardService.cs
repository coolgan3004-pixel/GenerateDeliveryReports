using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace GenerateDeliveryReports.Data.Services;

public class SprintDashboardService
{
    private readonly AppSettings _settings;
    private readonly ILogger<SprintDashboardService> _logger;

    public SprintDashboardService(IOptions<AppSettings> options, ILogger<SprintDashboardService> logger)
    {
        _settings = options.Value;
        _logger = logger;
    }

    public (bool exists, DateTime? lastUpdated) GetDashboardInfo()
    {
        var path = _settings.ResolvedSprintDashboardHtmlPath;
        _logger.LogInformation("SprintDashboard: HtmlPath='{Path}', FileExists={Exists}", path, File.Exists(path));
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return (false, null);
        return (true, File.GetLastWriteTime(path));
    }

    /// <summary>Whether enough config is present to generate the dashboard -- a metrics root
    /// folder to read workbooks from, plus an output path to write the HTML to.</summary>
    public bool IsScriptConfigured
    {
        get
        {
            var configured = !string.IsNullOrWhiteSpace(_settings.SprintDashboardHtmlPath)
                && !string.IsNullOrWhiteSpace(_settings.MetricsFolder)
                && !string.IsNullOrWhiteSpace(_settings.ResolvedOneDriveLocation);
            _logger.LogInformation("SprintDashboard: IsConfigured={Configured}", configured);
            return configured;
        }
    }

    /// <summary>
    /// Generates the dashboard natively in C# (see <see cref="SprintBriefGenerator"/>) -- this used
    /// to shell out to a Python script, which doesn't work on Azure App Service (no Python runtime
    /// there). Kept the method name so the calling Razor page didn't need to change.
    /// </summary>
    public async Task<(bool success, string message)> RunScriptAsync()
    {
        var outputPath = _settings.ResolvedSprintDashboardHtmlPath;
        if (string.IsNullOrWhiteSpace(outputPath))
            return (false, "Sprint dashboard output path (SprintDashboardHtmlPath) not configured.");

        var root = Path.Combine(_settings.ResolvedOneDriveLocation, _settings.MetricsFolder);
        if (!Directory.Exists(root))
            return (false, $"Metrics folder not found: {root}");

        try
        {
            var (success, message) = await SprintBriefGenerator.GenerateAsync(root, outputPath);
            if (!success)
                _logger.LogError("Sprint dashboard generation failed: {Message}", message);
            return (success, message);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to generate sprint dashboard.");
            return (false, ex.Message);
        }
    }
}
