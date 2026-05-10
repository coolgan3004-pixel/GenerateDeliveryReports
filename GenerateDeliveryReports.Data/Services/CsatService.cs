using GenerateDeliveryReports.Data.Concrete;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Options;
using OfficeOpenXml;

namespace GenerateDeliveryReports.Data.Services;

public class CsatService
{
    private readonly AppSettings _settings;

    public CsatService(IOptions<AppSettings> options)
    {
        _settings = options.Value;
    }

    public IEnumerable<string> GetClientNames() =>
        _settings.CSAT.Clients.Select(c => c.ClientName);

    public IEnumerable<string> GetSheetNames(string clientName)
    {
        var client = _settings.CSAT.Clients
            .FirstOrDefault(c => c.ClientName == clientName);
        return client?.SurveySheets ?? [];
    }

    private CsatClient? GetClient(string clientName) =>
        _settings.CSAT.Clients.FirstOrDefault(c => c.ClientName == clientName);

    public string GetDefaultFromEmail() => _settings.CSAT.FromEmailAddress;
    public string GetClientEmail(string clientName) => GetClient(clientName)?.ClientEmailAddress ?? string.Empty;
    public string GetClientSubject(string clientName) => GetClient(clientName)?.ClientEmailSubject ?? string.Empty;
    public string GetClientEmailBody(string clientName) => GetClient(clientName)?.ClientEmailBody ?? string.Empty;

    private string ResolveExcelPath(string clientName)
    {
        var client = GetClient(clientName)
            ?? throw new InvalidOperationException($"Client '{clientName}' not found.");

        var folder = Path.Combine(_settings.OneDriveLocation, _settings.CSAT.CSATFolder.TrimStart('\\'));
        return Path.GetFullPath(Path.Combine(folder, client.ClientSurveyFilePath));
    }

    /// <summary>
    /// Generates a PDF for the specified sheet via the HTML intermediary and returns
    /// a web-accessible URL (e.g. /downloads/SheetName.pdf).
    /// </summary>
    public string GenerateSheetPdf(string clientName, string sheetName)
    {
        var excelPath = ResolveExcelPath(clientName);
        if (!File.Exists(excelPath))
            throw new FileNotFoundException($"CSAT file not found: {excelPath}");

        var client = GetClient(clientName)
            ?? throw new InvalidOperationException($"Client '{clientName}' not found.");
        var outputDir = _settings.TempPath;

        int startCol = client.StartColumn > 0 ? client.StartColumn : 2;
        int endCol = client.EndColumn > 0 ? client.EndColumn : 7;
    

       

        using var wrapper = new ExcelWrapper();
        wrapper.Open(excelPath);
        var paths = wrapper.GeneratePdfFileFromWorkSheets(outputDir, [sheetName], startCol, endCol);

        var pdfPath = paths.FirstOrDefault()
            ?? throw new InvalidOperationException($"PDF generation failed for sheet '{sheetName}'.");

        return $"/downloads/{Path.GetFileName(pdfPath)}";
    }
}
