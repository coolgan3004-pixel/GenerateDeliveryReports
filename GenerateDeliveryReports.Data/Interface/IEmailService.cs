using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Data.Interface;

public interface IEmailService
{
    Task SendAsync(EmailParameters parameters);
}
