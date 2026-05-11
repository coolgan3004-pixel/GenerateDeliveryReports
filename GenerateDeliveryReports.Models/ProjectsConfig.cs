namespace GenerateDeliveryReports.Models;

public class AppSettings
{
    public string CommonFolderPath { get; set; } = string.Empty;

    /// <summary>Folder where generated PDFs and charts are written and served from /downloads.</summary>
    public string TempPath
    {
        get
        {
            var directory = string.IsNullOrWhiteSpace(CommonFolderPath)
                ? Path.Combine(AppContext.BaseDirectory, "wwwroot", "downloads")
                : Path.Combine(CommonFolderPath, "downloads");
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            return directory;
        }
    }

    /// <summary>Folder where rolling log files are written.</summary>
    public string LogFilesPath
    {
        get
        {
            var directory = string.IsNullOrWhiteSpace(CommonFolderPath)
                ? Path.Combine(AppContext.BaseDirectory, "LogFiles")
                : Path.Combine(CommonFolderPath, "LogFiles");
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
    public CsatConfig CSAT { get; set; } = new();
    public EmailSetting EmailSettings { get; set; } = new();
    public string PMOEmailContent { get; set; } = string.Empty;
    public int WorkerIntervalMinutes { get; set; }
}
