namespace GenerateDeliveryReports.Models;

public static class ErrorCodes
{
    // File & Path errors (1xx)
    public const string ERR_100 = "ERR_100";
    public const string ERR_101 = "ERR_101";
    public const string ERR_102 = "ERR_102";
    public const string ERR_103 = "ERR_103";

    // Excel / Data errors (2xx)
    public const string ERR_200 = "ERR_200";
    public const string ERR_201 = "ERR_201";
    public const string ERR_202 = "ERR_202";
    public const string ERR_203 = "ERR_203";

    // Report generation errors (3xx)
    public const string ERR_300 = "ERR_300";
    public const string ERR_301 = "ERR_301";
    public const string ERR_302 = "ERR_302";
    public const string ERR_303 = "ERR_303";

    // Configuration errors (4xx)
    public const string ERR_400 = "ERR_400";
    public const string ERR_401 = "ERR_401";

    private static readonly Dictionary<string, string> Messages = new()
    {
        // File & Path errors
        { ERR_100, "Data file not found at the configured path: {0}" },
        { ERR_101, "No matching file found after searching for recently modified versions of: {0}" },
        { ERR_102, "Metrics sheet file not found: {0}" },
        { ERR_103, "PPT template file not found: {0}" },

        // Excel / Data errors
        { ERR_200, "No sprint data found in the Data sheet for project '{0}'." },
        { ERR_201, "Failed to read Excel file: {0}" },
        { ERR_202, "More than 1 record present in the data list." },
        { ERR_203, "Chart not found in the Scorecard sheet." },

        // Report generation errors
        { ERR_300, "More than 3 slides found in the PPT template." },
        { ERR_301, "Failed to generate presentation: {0}" },
        { ERR_302, "Failed to export chart image: {0}" },
        { ERR_303, "Chart image file not found for report generation: {0}" },

        // Configuration errors
        { ERR_400, "Project '{0}' not found in configuration." },
        { ERR_401, "OneDriveLocation is not configured in AppSettings." },
    };

    public static string GetMessage(string code, params object[] args)
    {
        if (Messages.TryGetValue(code, out var template))
            return $"[{code}] {string.Format(template, args)}";
        return $"[{code}] An unknown error occurred.";
    }
}
