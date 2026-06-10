using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Logging;

namespace GenerateDeliveryReports.Data.Services;

/// <summary>
/// Sends email via the Outlook desktop client using COM automation.
/// Outlook must be installed and configured on the machine running this service.
/// </summary>
public class OutlookEmailService : IEmailService
{
    private readonly ILogger<OutlookEmailService> _logger;

    public OutlookEmailService(ILogger<OutlookEmailService> logger)
    {
        _logger = logger;
    }

    public (bool success, string message) Send(EmailParameters parameters)
    {
        if (string.IsNullOrWhiteSpace(parameters.ToEmailAddress))
            return (false, "Recipient (To) address is required.");

        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException(
                    "Outlook is not installed or not registered on this machine.");

            dynamic outlook = Activator.CreateInstance(outlookType)!;
            dynamic mail    = outlook.CreateItem(0); // 0 = olMailItem

            // -- Recipients --
            // ToEmailAddress may contain multiple addresses separated by ; or ,
            var recipients = parameters.ToEmailAddress
                .Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            foreach (var recipient in recipients)
                mail.Recipients.Add(recipient);

            mail.Recipients.ResolveAll();

            // -- Subject and body --
            mail.Subject = parameters.Subject;

            if (parameters.IsHtmlBody)
                mail.HTMLBody = parameters.Body;
            else
                mail.Body = parameters.Body;

            // -- Sender account --
            // Walk the Outlook Session.Accounts collection to find the account whose
            // SMTP address matches FromEmailAddress, then set SendUsingAccount so the
            // email is sent from that specific account rather than the default.
            if (!string.IsNullOrWhiteSpace(parameters.FromEmailAddress))
            {
                try
                {
                    dynamic accounts = outlook.Session.Accounts;
                    for (int i = 1; i <= accounts.Count; i++)
                    {
                        dynamic acct = accounts[i];
                        string smtpAddress = acct.SmtpAddress?.ToString() ?? string.Empty;
                        if (string.Equals(smtpAddress, parameters.FromEmailAddress,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            mail.SendUsingAccount = acct;
                            _logger.LogInformation("OutlookEmailService: sending from account {Account}.", smtpAddress);
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Non-fatal: fall back to Outlook default account
                    _logger.LogWarning(ex,
                        "OutlookEmailService: could not resolve From account '{From}'. Falling back to default.",
                        parameters.FromEmailAddress);
                }
            }

            // -- Attachments --
            foreach (var path in parameters.Attachments)
            {
                if (File.Exists(path))
                    mail.Attachments.Add(path);
                else
                    _logger.LogWarning("OutlookEmailService: attachment not found, skipping: {Path}", path);
            }

            mail.Send();

            _logger.LogInformation("OutlookEmailService: email sent to {To} with subject '{Subject}'.",
                parameters.ToEmailAddress, parameters.Subject);

            return (true, $"Email sent successfully to {parameters.ToEmailAddress}.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "OutlookEmailService: failed to send email to {To}.", parameters.ToEmailAddress);
            return (false, $"Failed to send email: {ex.Message}");
        }
    }
}
