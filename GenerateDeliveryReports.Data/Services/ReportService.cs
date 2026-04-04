using GenerateDeliveryReports.Data.Settings;
using GenerateDeliveryReports.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace GenerateDeliveryReports.Data.Services;

public class ReportService : IReportService
{
    private readonly IWebHostEnvironment _env;
    private readonly IConfiguration _configuration;
    private readonly ILogger<ReportService> _logger;
    private readonly ProjectSettingsLoader _settingsLoader;

    public ReportService(
        IWebHostEnvironment env,
        IConfiguration configuration,
        ILogger<ReportService> logger,
        ProjectSettingsLoader settingsLoader)
    {
        _env = env;
        _configuration = configuration;
        _logger = logger;
        _settingsLoader = settingsLoader;
    }

    /// <inheritdoc />
    public Task<List<Project>> GetProjectsAsync()
    {
        return _settingsLoader.LoadProjectsAsync(_env.ContentRootPath);
    }

    /// <inheritdoc />
    public async Task<List<string>> ProcessDeliverySummaryFileAsync(string filePath)
    {
        // Resolve the full path: the downloads folder is configurable via
        // appsettings.json ("DownloadsFolder"). Defaults to the current user's
        // system Downloads folder so the app works out-of-the-box when run
        // locally.
        var downloadsFolder = _configuration["DownloadsFolder"]
            ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads");

        // Use only the file name portion to prevent path-traversal attacks.
        var safeFileName = Path.GetFileName(filePath);
        var fullPath = Path.Combine(downloadsFolder, safeFileName);

        if (!File.Exists(fullPath))
        {
            _logger.LogWarning(
                "Delivery summary file not found: {FullPath}", fullPath);
            return new List<string>();
        }

        var results = new List<string>();
        var lines = await File.ReadAllLinesAsync(fullPath);

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed))
                continue;

            // If the file is a CSV, split on commas and add each non-empty
            // field; otherwise treat the whole line as a single entry.
            if (trimmed.Contains(','))
            {
                var fields = trimmed.Split(',')
                    .Select(f => f.Trim())
                    .Where(f => !string.IsNullOrEmpty(f));
                results.AddRange(fields);
            }
            else
            {
                results.Add(trimmed);
            }
        }

        return results;
    }
}
