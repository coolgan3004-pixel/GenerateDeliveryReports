using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Data.Interface;

public interface IEmailService
{
    /// <summary>
    /// Sends an email using the configured provider.
    /// </summary>
    /// <param name="parameters">Recipient, sender, subject, body, attachments.</param>
    /// <returns>Success flag and a human-readable result message.</returns>
    (bool success, string message) Send(EmailParameters parameters);
}
