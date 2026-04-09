namespace GenerateDeliveryReports.Models;

public class EmailUser
{
    public string Name { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
}

public class EmailSetting
{
    public string Provider { get; set; } = string.Empty;
    public string SMTPServer { get; set; } = string.Empty;
    public int Port { get; set; }
    public bool EnableSsl { get; set; } = true;
    public string UserName { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public string FromEmailAddress { get; set; } = string.Empty;
    public List<EmailUser> Users { get; set; } = [];
}
