namespace GenerateDeliveryReports.Models;

public class AppSettings
{
    public string TempPath
    {
        get
        {
            var directory = Path.Combine(AppContext.BaseDirectory, "wwwroot", "downloads");
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            return directory;
        }
    }

    public string OneDriveLocation { get; set; } = string.Empty;
    public string ReportAndDataFolder { get; set; } = string.Empty;
    public string MetricsFolder { get; set; } = string.Empty;
    public string SprintMetricsReportTemplatePath { get; set; }= string.Empty;
    public string WorkerSummaryFilePath { get; set; } = string.Empty;
    public List<Project> Projects { get; set; } = [];
    public EmailSetting EmailSettings { get; set; } = new();
    public string PMOEmailContent { get; set; } = string.Empty;
    public int WorkerIntervalMinutes { get; set; }
}
