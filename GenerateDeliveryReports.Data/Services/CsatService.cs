using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Options;
using Spire.Xls;

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
        return client?.SheetNames ?? [];
    }

    public string GetDefaultFromEmail() => _settings.CSAT.FromEmailAddress;
    public string GetDefaultSubject() => _settings.CSAT.Subject;

    public string GetClientEmail(string clientName) =>
        _settings.CSAT.Clients
            .FirstOrDefault(c => c.ClientName == clientName)?.ClientEmailAddress ?? string.Empty;

    private string ResolveExcelPath(string clientName)
    {
        var client = _settings.CSAT.Clients
            .FirstOrDefault(c => c.ClientName == clientName)
            ?? throw new InvalidOperationException($"Client '{clientName}' not found.");

        var folder = Path.Combine(_settings.OneDriveLocation, _settings.CSAT.CSATFolder.TrimStart('\\'));
        return Path.GetFullPath(Path.Combine(folder, client.CSATFileName));
    }

    /// <summary>
    /// Generates a PDF for the specified sheet and returns a web-accessible relative URL.
    /// Returns the path relative to wwwroot (e.g. /downloads/csat_SheetName.pdf).
    /// </summary>
    public string GenerateSheetPdf(string clientName, string sheetName)
    {
        var excelPath = ResolveExcelPath(clientName);
        if (!File.Exists(excelPath))
            throw new FileNotFoundException($"CSAT file not found: {excelPath}");

        var safeSheet = string.Concat(sheetName.Split(Path.GetInvalidFileNameChars()));
        var fileName = $"csat_{safeSheet}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
        var outputDir = Path.Combine(AppContext.BaseDirectory, "wwwroot", "downloads");
        Directory.CreateDirectory(outputDir);
        var outputPath = Path.Combine(outputDir, fileName);

        var workbook = new Workbook();
        workbook.LoadFromFile(excelPath);

        // Validate the sheet exists
        var ws = workbook.Worksheets[sheetName]
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found in workbook.");

        // Remove all other sheets so only the target sheet is exported to PDF
        for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
        {
            if (workbook.Worksheets[i].Name != sheetName)
                workbook.Worksheets.RemoveAt(i);
        }

        workbook.SaveToFile(outputPath, Spire.Xls.FileFormat.PDF);

        return $"/downloads/{fileName}";
    }
}
