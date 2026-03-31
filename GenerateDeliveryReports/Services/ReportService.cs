using System.Text.Json;
using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Services;

public class ReportService : IReportService
{
    private readonly IWebHostEnvironment _env;
    private readonly IConfiguration _configuration;
    private readonly ILogger<ReportService> _logger;

    public ReportService(
        IWebHostEnvironment env,
        IConfiguration configuration,
        ILogger<ReportService> logger)
    {
        _env = env;
        _configuration = configuration;
        _logger = logger;
    }

    /// <inheritdoc />
    public async Task<List<Project>> GetProjectsAsync()
    {
        var configPath = Path.Combine(_env.ContentRootPath, "Data", "projects.json");

        if (!File.Exists(configPath))
        {
            _logger.LogWarning("projects.json not found at {Path}", configPath);
            return new List<Project>();
        }

        await using var stream = File.OpenRead(configPath);
        var config = await JsonSerializer.DeserializeAsync<ProjectsConfig>(
            stream,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        return config?.Projects ?? new List<Project>();
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
