namespace GenerateDeliveryReports.Models;

public class CsatClient
{
    public string CSATType { get; set; } = string.Empty;
    public string ClientName { get; set; } = string.Empty;
    public string ClientEmailAddress { get; set; } = string.Empty;
    public string ClientEmailSubject { get; set; } = string.Empty;
    public string ClientEmailBody { get; set; } = string.Empty;
    public string ClientSurveyFilePath { get; set; } = string.Empty;
    public List<string> SurveySheets { get; set; } = [];
    /// <summary>First column to include in the PDF export (1-based). Defaults to 1.</summary>
    public int StartColumn { get; set; } = 1;
    /// <summary>Last column to include in the PDF export (1-based). 0 = auto-detect from sheet dimension.</summary>
    public int EndColumn { get; set; } = 0;
}

public class CsatConfig
{
    public string CSATFolder { get; set; } = string.Empty;
    public string FromEmailAddress { get; set; } = string.Empty;
    public List<CsatClient> Clients { get; set; } = [];
}
