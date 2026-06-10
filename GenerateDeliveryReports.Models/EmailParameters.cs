namespace GenerateDeliveryReports.Models;

public class EmailParameters
{
    public string ToEmailAddress { get; set; } = string.Empty;
    public string FromEmailAddress { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    /// <summary>True to set HTMLBody; false (default) to set plain-text Body.</summary>
    public bool IsHtmlBody { get; set; } = false;
    public List<string> Attachments { get; set; } = [];
}
