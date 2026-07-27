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

    /// <summary>
    /// Resolves <see cref="OneDriveLocation"/> to an absolute path. Locally, this is normally an
    /// absolute (or UNC) path to a real OneDrive sync folder, so it's used as-is. On Azure App
    /// Service, there is no OneDrive sync -- the xlsx files are bundled into the deployment
    /// instead, so the config value is a relative path (e.g. "CommonFiles\Files"), resolved here
    /// against the app's own base directory.
    /// </summary>
    public string ResolvedOneDriveLocation =>
        string.IsNullOrWhiteSpace(OneDriveLocation) || Path.IsPathRooted(OneDriveLocation)
            ? OneDriveLocation
            : Path.Combine(AppContext.BaseDirectory, OneDriveLocation);
    public string ReportAndDataFolder { get; set; } = string.Empty;
    public string MetricsFolder { get; set; } = string.Empty;
    public string SprintMetricsReportTemplatePath { get; set; }= string.Empty;
    public string WorkerSummaryFilePath { get; set; } = string.Empty;
    public string SprintDashboardHtmlPath { get; set; } = string.Empty;
    public string SprintReportStatusJsonPath { get; set; } = string.Empty;
    public string SprintDashboardScriptPath { get; set; } = string.Empty;
    public string PythonExePath { get; set; } = "python";
    public List<Project> Projects { get; set; } = [];
    public CsatConfig CSAT { get; set; } = new();
    public EmailSetting EmailSettings { get; set; } = new();
    public string PMOEmailContent { get; set; } = string.Empty;
    public int WorkerIntervalMinutes { get; set; }

    public string AnthropicApiKey { get; set; } = string.Empty;
    public string BriefingModel { get; set; } = "claude-sonnet-5";
    public int BriefingDailyTriggerHour { get; set; } = 8;
    public int BriefingMaxWebSearches { get; set; } = 8;
    public int BriefingMaxTokens { get; set; } = 8000;
    public string BriefingArchiveFolder { get; set; } = string.Empty;
}
