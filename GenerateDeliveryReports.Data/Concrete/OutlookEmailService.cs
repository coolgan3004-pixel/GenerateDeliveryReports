using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Logging;

namespace GenerateDeliveryReports.Data.Concrete;

public class OutlookEmailService : IEmailService
{
    private readonly ILogger<OutlookEmailService> _logger;

    public OutlookEmailService(ILogger<OutlookEmailService> logger)
    {
        _logger = logger;
    }

    public Task SendAsync(EmailParameters parameters)
    {
        var outlookType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Outlook is not installed or not registered on this machine.");

        dynamic outlook = Activator.CreateInstance(outlookType)!;
        dynamic mail = outlook.CreateItem(0); // 0 = olMailItem

        mail.Subject = parameters.Subject;
        mail.HTMLBody = parameters.Body;

        // Support semicolon- or comma-separated recipient lists
        var recipients = parameters.ToEmailAddress
            .Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        if (recipients.Length == 0)
            throw new InvalidOperationException("No recipients specified.");

        foreach (var recipient in recipients)
            mail.Recipients.Add(recipient);

        mail.Recipients.ResolveAll();

        foreach (var attachment in parameters.Attachments)
        {
            if (File.Exists(attachment))
                mail.Attachments.Add(attachment);
            else
                _logger.LogWarning("Attachment not found, skipping: {Path}", attachment);
        }

        mail.Send();

        _logger.LogInformation(
            "Email '{Subject}' sent via Outlook to {Recipients}",
            parameters.Subject, parameters.ToEmailAddress);

        return Task.CompletedTask;
    }
}
