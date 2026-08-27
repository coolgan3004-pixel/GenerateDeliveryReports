using MimeKit;
using System.IO;

namespace GenerateDeliveryReports.Services;

public class EmailMessageGenerator
{
    public byte[] GenerateOutlookMessage(
        string fromEmail,
        string toEmail,
        string subject,
        string htmlBody,
        string? ccEmails = null,
        IEnumerable<string>? attachmentPaths = null)
    {
        // Create MIME message
        var message = new MimeMessage();

        message.From.Add(new MailboxAddress("", fromEmail));
        message.To.Add(new MailboxAddress("", toEmail));
        message.Subject = subject;

        // Add CC recipients if provided
        if (!string.IsNullOrEmpty(ccEmails))
        {
            var ccList = ccEmails.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var cc in ccList.Select(c => c.Trim()))
            {
                message.Cc.Add(new MailboxAddress("", cc));
            }
        }

        // Create HTML body with attachments
        var builder = new BodyBuilder { HtmlBody = htmlBody };

        // Add attachments if provided
        if (attachmentPaths != null)
        {
            foreach (var filePath in attachmentPaths)
            {
                if (File.Exists(filePath))
                {
                    builder.Attachments.Add(filePath);
                }
            }
        }

        message.Body = builder.ToMessageBody();

        // Convert to EML (MIME format) as byte array
        using (var memoryStream = new MemoryStream())
        {
            message.WriteTo(memoryStream);
            return memoryStream.ToArray();
        }
    }
}
