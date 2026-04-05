namespace GenerateDeliveryReports.Models;

public class EmailParameters
{
    public string ToEmailAddress { get; set; } = string.Empty;
    public string FromEmailAddress { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public List<string> Attachments { get; set; } = [];
    public string Body { get; set; } = string.Empty;
}
