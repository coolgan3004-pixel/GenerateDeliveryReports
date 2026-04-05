namespace GenerateDeliveryReports.Models;

public class Project
{
    public string ProjectName { get; set; } = string.Empty;
    public string[] MetricsSheetPath { get; set; } = [];
    public string DataFileName { get; set; } = string.Empty;
    public string ProjectFolderOneDriveLink { get; set; } = string.Empty;
}
