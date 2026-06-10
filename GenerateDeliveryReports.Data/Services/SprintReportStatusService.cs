using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Options;
using System.Text.Json;

namespace GenerateDeliveryReports.Data.Services;

public class SprintReportStatusService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly AppSettings _settings;

    public SprintReportStatusService(IOptions<AppSettings> options)
    {
        _settings = options.Value;
    }

    /// <summary>
    /// Returns (report, lastModified, errorMessage).
    /// report is null when the file does not exist or cannot be read.
    /// </summary>
    public (SprintStatusReport? Report, DateTime? LastModified, string? Error) Load()
    {
        var path = ResolveJsonPath();

        if (!File.Exists(path))
            return (null, null, null);

        try
        {
            var lastModified = File.GetLastWriteTime(path);
            var json         = File.ReadAllText(path);
            var report       = JsonSerializer.Deserialize<SprintStatusReport>(json, JsonOptions)
                               ?? new SprintStatusReport();
            return (report, lastModified, null);
        }
        catch (Exception ex)
        {
            return (null, null, $"Failed to read status file: {ex.Message}");
        }
    }

    public string ResolveJsonPath()
    {
        return string.IsNullOrWhiteSpace(_settings.SprintReportStatusJsonPath)
            ? Path.Combine(AppContext.BaseDirectory, "sprint-report-status.json")
            : _settings.SprintReportStatusJsonPath;
    }
}
