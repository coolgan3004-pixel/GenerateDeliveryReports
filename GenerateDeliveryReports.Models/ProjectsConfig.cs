namespace GenerateDeliveryReports.Models;

public class AppSettings
{
    public string CommonFolderPath { get; set; } = string.Empty;

    /// <summary>
    /// Base directory for files the app WRITES at runtime (logs, the SQLite identity DB,
    /// generated downloads, the briefing archive) when <see cref="CommonFolderPath"/> isn't set.
    /// Azure App Service's "Run From Package" zip deploy mounts the app's own folder
    /// (AppContext.BaseDirectory / wwwroot) READ-ONLY, so writing there throws -- confusingly,
    /// as a FileNotFoundException from Directory.CreateDirectory rather than an access-denied
    /// error. Azure always sets a HOME environment variable pointing at persistent, writable
    /// storage outside wwwroot (normally D:\home); local/on-prem runs never have HOME set, so
    /// checking for it here only changes behavior on Azure.
    /// </summary>
    private static string WritableBase =>
        Environment.GetEnvironmentVariable("HOME") is { Length: > 0 } home
            ? Path.Combine(home, "data", "GenerateDeliveryReports")
            : AppContext.BaseDirectory;

    /// <summary>Folder where generated PDFs and charts are written and served from /downloads.</summary>
    public string TempPath
    {
        get
        {
            var directory = string.IsNullOrWhiteSpace(CommonFolderPath)
                ? Path.Combine(WritableBase, "wwwroot", "downloads")
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
                ? Path.Combine(WritableBase, "LogFiles")
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
                ? Path.Combine(WritableBase, "App_Data")
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
    /// against the app's own base directory. This is READ-only access to the bundled files, so
    /// AppContext.BaseDirectory (wwwroot) is correct here even under Run From Package -- unlike
    /// the write paths above, reading from a read-only-mounted wwwroot works fine.
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

    /// <summary>Same relative/absolute resolution as <see cref="ResolvedBriefingArchiveFolder"/>, but for
    /// the dashboard's output HTML file itself: an absolute path is used as-is; a relative one resolves
    /// against WritableBase so the app can both read AND write it on Azure App Service (unlike
    /// <see cref="ResolvedOneDriveLocation"/>, which only ever needs to be read).</summary>
    public string ResolvedSprintDashboardHtmlPath =>
        string.IsNullOrWhiteSpace(SprintDashboardHtmlPath) || Path.IsPathRooted(SprintDashboardHtmlPath)
            ? SprintDashboardHtmlPath
            : Path.Combine(WritableBase, SprintDashboardHtmlPath);
    public string SprintReportStatusJsonPath { get; set; } = string.Empty;
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
    /// CommonFolderPath (or WritableBase if that's blank too).
    /// </summary>
    public string ResolvedBriefingArchiveFolder
    {
        get
        {
            var directory = string.IsNullOrWhiteSpace(BriefingArchiveFolder)
                ? (string.IsNullOrWhiteSpace(CommonFolderPath)
                    ? Path.Combine(WritableBase, "BriefingArchive")
                    : Path.Combine(CommonFolderPath, "BriefingArchive"))
                : (Path.IsPathRooted(BriefingArchiveFolder)
                    ? BriefingArchiveFolder
                    : Path.Combine(string.IsNullOrWhiteSpace(CommonFolderPath) ? WritableBase : CommonFolderPath, BriefingArchiveFolder));
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
            return directory;
        }
    }
}
