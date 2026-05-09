namespace GenerateDeliveryReports.Models;

public class CsatClient
{
    public string ClientName { get; set; } = string.Empty;
    public string ClientEmailAddress { get; set; } = string.Empty;
    public string CSATFileName { get; set; } = string.Empty;
    public List<string> SheetNames { get; set; } = [];
}

public class CsatConfig
{
    public string CSATFolder { get; set; } = string.Empty;
    public string FromEmailAddress { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public List<CsatClient> Clients { get; set; } = [];
}
