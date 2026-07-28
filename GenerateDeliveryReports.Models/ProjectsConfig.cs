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

    /// <summary>SQLite connection string for the local sign-in accounts database (ASP.NET Core Identity).</summary>
    public string IdentityConnectionString
    {
        get
        {
            var directory = string.IsNullOrWhiteSpace(CommonFolderPath)
                ? Path.Combine(AppContext.BaseDirectory, "App_Data")
                : Path.Combine(CommonFolderPath, "App_Data");
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            return $"Data Source={Path.Combine(directory, "identity.db")}";
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

    /// <summary>Same relative/absolute resolution as <see cref="ResolvedOneDriveLocation"/>, for the
    /// PPTX template when it's bundled into the deployment instead of read from a fixed disk path.</summary>
    public string ResolvedSprintMetricsReportTemplatePath =>
        string.IsNullOrWhiteSpace(SprintMetricsReportTemplatePath) || Path.IsPathRooted(SprintMetricsReportTemplatePath)
            ? SprintMetricsReportTemplatePath
            : Path.Combine(AppContext.BaseDirectory, SprintMetricsReportTemplatePath);
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

    /// <summary>
    /// Resolves <see cref="BriefingArchiveFolder"/> the same way as <see cref="TempPath"/>/<see cref="LogFilesPath"/>:
    /// an absolute path is used as-is; a relative or empty value falls back to a folder under
    /// CommonFolderPath (or the app's own base directory if that's blank too). Without this, a
    /// relative value would resolve against the process's working directory, which isn't
    /// guaranteed to be the app's own folder on every host.
    /// </summary>
    public string ResolvedBriefingArchiveFolder
    {
        get
        {
            var directory = string.IsNullOrWhiteSpace(BriefingArchiveFolder)
                ? (string.IsNullOrWhiteSpace(CommonFolderPath)
                    ? Path.Combine(AppContext.BaseDirectory, "BriefingArchive")
                    : Path.Combine(CommonFolderPath, "BriefingArchive"))
                : (Path.IsPathRooted(BriefingArchiveFolder)
                    ? BriefingArchiveFolder
                    : Path.Combine(string.IsNullOrWhiteSpace(CommonFolderPath) ? AppContext.BaseDirectory : CommonFolderPath, BriefingArchiveFolder));
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            return directory;
        }
    }
}
