using GenerateDeliveryReports.Data.Services;
using GenerateDeliveryReports.Services;
using Microsoft.AspNetCore.Mvc;

namespace GenerateDeliveryReports.Controllers;

[ApiController]
[Route("api/csat")]
public class CsatEmailController : ControllerBase
{
    private readonly CsatService _csatService;
    private readonly EmailMessageGenerator _emailGenerator;

    public CsatEmailController(CsatService csatService, EmailMessageGenerator emailGenerator)
    {
        _csatService = csatService;
        _emailGenerator = emailGenerator;
    }

    /// <summary>
    /// Download CSAT email as .eml file (Outlook/any email client compatible)
    /// </summary>
    [HttpGet("download-email")]
    public IActionResult DownloadEmail(
        [FromQuery] string clientName,
        [FromQuery] string from,
        [FromQuery] string to,
        [FromQuery] string subject,
        [FromQuery] string body,
        [FromQuery(Name = "cc")] string? cc = null,
        [FromQuery(Name = "attachments")] string? attachments = null)
    {
        try
        {
            // Parse attachment paths (semicolon-separated)
            var attachmentPaths = attachments?.Split(';', StringSplitOptions.RemoveEmptyEntries) ?? [];

            // Generate email message with attachments
            byte[] emlBytes = _emailGenerator.GenerateOutlookMessage(from, to, subject, body, cc, attachmentPaths);

            // Return as downloadable EML file
            return File(emlBytes, "message/rfc822", $"{clientName}_{DateTime.Now:yyyyMMdd_HHmmss}.eml");
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }
}
