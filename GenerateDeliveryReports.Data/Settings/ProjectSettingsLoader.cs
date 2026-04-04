using System.Text.Json;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Logging;

namespace GenerateDeliveryReports.Data.Settings;

/// <summary>
/// Responsible for loading application/project settings from the
/// projectsettings.json configuration file.
/// </summary>
public class ProjectSettingsLoader
{
    private readonly ILogger<ProjectSettingsLoader> _logger;

    public ProjectSettingsLoader(ILogger<ProjectSettingsLoader> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Loads the list of projects from the projectsettings.json file located
    /// in the given base directory.
    /// </summary>
    /// <param name="contentRootPath">
    /// The application's content root path (typically <c>IWebHostEnvironment.ContentRootPath</c>).
    /// </param>
    public async Task<List<Project>> LoadProjectsAsync(string contentRootPath)
    {
        var configPath = Path.Combine(contentRootPath, "Data", "projectsettings.json");

        if (!File.Exists(configPath))
        {
            _logger.LogWarning("projectsettings.json not found at {Path}", configPath);
            return new List<Project>();
        }

        await using var stream = File.OpenRead(configPath);
        var config = await JsonSerializer.DeserializeAsync<ProjectsConfig>(
            stream,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        return config?.Projects ?? new List<Project>();
    }
}
