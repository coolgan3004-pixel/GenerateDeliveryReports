using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Options;
using System.Text.Json;

namespace GenerateDeliveryReports.Data.Services;

public class SprintReportsAvailabilityService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly AppSettings _settings;

    public SprintReportsAvailabilityService(IOptions<AppSettings> options)
    {
        _settings = options.Value;
    }

    /// <summary>
    /// Returns (matrix, lastModified, errorMessage).
    /// matrix is null when the file does not exist or cannot be read.
    /// </summary>
    public (SprintReportsAvailabilityMatrix? Matrix, DateTime? LastModified, string? Error) Load()
    {
        var path = ResolveJsonPath();

        if (!File.Exists(path))
        {
            // Log the path being attempted (for debugging)
            var errorMsg = $"File not found at: {path}. BaseDirectory: {AppContext.BaseDirectory}. Configured: {_settings.SprintReportsAvailabilityJSONFileName}";
            return (null, null, errorMsg);
        }

        try
        {
            var lastModified = File.GetLastWriteTime(path);
            var json         = File.ReadAllText(path);
            var matrix       = JsonSerializer.Deserialize<SprintReportsAvailabilityMatrix>(json, JsonOptions)
                               ?? new SprintReportsAvailabilityMatrix();
            return (matrix, lastModified, null);
        }
        catch (Exception ex)
        {
            return (null, null, $"Failed to read availability matrix file: {ex.Message}");
        }
    }

    public string ResolveJsonPath()
    {
        var fileName = _settings.SprintReportsAvailabilityJSONFileName ?? "SprintReportsAvailability.json";

        // If it's a full path, use as-is; otherwise combine with BaseDirectory
        if (string.IsNullOrEmpty(Path.GetDirectoryName(fileName)) || !Path.IsPathRooted(fileName))
        {
            return Path.Combine(AppContext.BaseDirectory, fileName);
        }

        return fileName;
    }
}
