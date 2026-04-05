namespace GenerateDeliveryReports.Models;

public class EmailSetting
{
    public string Provider { get; set; } = string.Empty;
    public string SMTPServer { get; set; } = string.Empty;
    public int Port { get; set; }
    public string UserName { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string FromEmailAddress { get; set; } = string.Empty;
}
