using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Services;

public interface IReportService
{
    /// <summary>
    /// Loads the list of projects from the projects.json configuration file.
    /// </summary>
    Task<List<Project>> GetProjectsAsync();

    /// <summary>
    /// Processes the delivery summary file for the given project and returns
    /// a list of string items (e.g. sprint names, release entries) to populate
    /// the second dropdown.
    /// </summary>
    /// <param name="filePath">
    /// File name (or relative path) of the delivery summary file, as specified
    /// in the projects.json configuration.
    /// </param>
    Task<List<string>> ProcessDeliverySummaryFileAsync(string filePath);
}
