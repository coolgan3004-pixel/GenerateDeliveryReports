using GenerateDeliveryReports.Data.Extensions;
using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using Spire.Xls;
using System.Drawing;
using System.Text.RegularExpressions;

namespace GenerateDeliveryReports.Data.Concrete;

public class DataProcessor : IDataProcessor
{
    private readonly AppSettings _appSettings;
    private readonly ILogger<DataProcessor> _logger;

    public DataProcessor(IOptions<AppSettings> appSettings, ILogger<DataProcessor> logger)
    {
        _appSettings = appSettings.Value;
        _logger = logger;
    }

    public string GetEmailContent(string projectName, string sprintName)
    {
        string emailContent = _appSettings.PMOEmailContent;
        string oneDriveURL = _appSettings.Projects
            .Where(x => x.ProjectName == projectName).SingleOrDefault()?.ProjectFolderOneDriveLink ?? "";

        var fileName = $"GlobalPayments -{projectName}-DeliveryQualitySummaryReport -{sprintName}";
        emailContent = emailContent.Replace("##", fileName);
        emailContent = emailContent.Replace("#Link#", $"{oneDriveURL}{fileName}");

        return emailContent;
    }

    public IEnumerable<string> GetProjectNames()
    {
        return _appSettings.Projects.Select(x => x.ProjectName).ToList();
    }

    public IEnumerable<string> GetSprintNames(string projectName)
    {
        using var excelWrapper = new ExcelWrapper();
        try
        {
            var project = _appSettings.Projects.FirstOrDefault(x => x.ProjectName == projectName);
            if (project == null)
                throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_400, projectName));

            var path = Path.Combine(_appSettings.OneDriveLocation, _appSettings.ReportAndDataFolder, project.DataFileName);
            _logger.LogInformation("GetSprintNames - Resolved path: {Path}", path);

            if (!File.Exists(path))
            {
                _logger.LogWarning("GetSprintNames - File.Exists returned false for: {Path}", path);
                throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_100, path));
            }

            var recentFile = path.GetRecentlyModifiedSimilarFile();
            if (recentFile == null)
            {
                _logger.LogWarning("GetSprintNames - GetRecentlyModifiedSimilarFile returned null for: {Path}", path);
                throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_101, path));
            }

            _logger.LogInformation("GetSprintNames - Opening file: {File}", recentFile.FullName);
            excelWrapper.Open(recentFile.FullName);

            var reportData = GetReportDataForProject(projectName).ToList();
            _logger.LogInformation("GetSprintNames - Report data rows: {Count}", reportData.Count);

            var sprintNames = reportData.Select(x => x.Sprint).Skip(1).ToArray();
            _logger.LogInformation("GetSprintNames - Sprint names found: {Count}", sprintNames.Length);

            if (sprintNames.Length == 0)
                throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_200, projectName));

            return sprintNames;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "GetSprintNames failed: {Message}", ex.Message);
            throw;
        }
    }

    public IEnumerable<SprintInfo> GetSprintNamesWithDate(string projectName)
    {
        var sprintNames = GetSprintNames(projectName);
        var project = _appSettings.Projects.Single(x => x.ProjectName == projectName);
        var projectDir = Path.GetDirectoryName(
            Path.Combine(_appSettings.OneDriveLocation, _appSettings.ReportAndDataFolder, project.DataFileName))!;

        var result = new List<SprintInfo>();
        foreach (var sprintName in sprintNames)
        {
            _logger.LogInformation("Processing sprint name: {SprintName} , Project:{ProjectName}", sprintName, projectName); 
            if (!string.IsNullOrEmpty(sprintName) && sprintName.IndexOf("2026", StringComparison.Ordinal) > 0)
            {

                var info = new SprintInfo
                {
                    ProjectName = projectName,
                    SprintName = sprintName
                };

                // Parse dates from sprint name format: "Sprint X (DD-MMM-YYYY to DD-MMM-YYYY)" or numeric variants
                var match = System.Text.RegularExpressions.Regex.Match(
                    sprintName,
                    @"\((\d{1,2}[\/\-\.]\w{1,9}[\/\-\.]\d{2,4})\s*to\s*(\d{1,2}[\/\-\.]\w{1,9}[\/\-\.]\d{2,4})\)");

                if (match.Success)
                {
                    var formats = new[]
                    {
                        "d-MMM-yyyy", "dd-MMM-yyyy", "MMM-d-yyyy", "MMM-dd-yyyy",
                        "d-MMMM-yyyy", "dd-MMMM-yyyy",
                        "dd/MM/yyyy", "d/M/yyyy", "MM/dd/yyyy", "M/d/yyyy",
                        "dd-MM-yyyy", "d-M-yyyy", "dd.MM.yyyy"
                    };
                    if (DateTime.TryParseExact(match.Groups[1].Value, formats,
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out var start))
                        info.SprintStartDate = start;

                    if (DateTime.TryParseExact(match.Groups[2].Value, formats,
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out var end))
                        info.SprintEndDate = end;
                }
                else
                {
                    _logger.LogWarning("GetSprintNamesWithDate - Could not parse dates from sprint name: {SprintName}", sprintName);
                }

                var sprintNameFormatted = sprintName.Contains('(')
                    ? sprintName[..sprintName.IndexOf('(')].Trim()
                    : sprintName.Trim();

                info.OutputPPTPath = Path.Combine(projectDir,
                    $"GlobalPayments-{projectName}-DeliveryQualitySummaryReport-{sprintNameFormatted}.pptx");

                result.Add(info);
            }
        }

        return result;
    }

    public SprintMetrics? GetSprintMetrics(string projectName, string sprintName)
    {
        List<DashboardData> datas = new();
        var colVarPair = new Dictionary<string, string>
        {
            { "Sprint #", "SprintName" },
            { "Sprint Start Date", "SprintStartDate" },
            { "Sprint End Date", "SprintEndDate" },
            { "Assigned", "Assigned" },
            { "Completed", "Completed" },
            { "Remarks", "Remarks" }
        };

        

        var reportData = GetReportDataForProject(projectName);
        SprintMetrics? sprintMetrics = reportData.Where(x => x.Sprint == sprintName).SingleOrDefault();
        
        var sprintNameFormatted = sprintName.Contains('(')
                ? sprintName[..sprintName.IndexOf('(')].Trim()
                : sprintName.Trim();

        /*if (sprintMetrics == null)
        {
            using var excelWrapper = new ExcelWrapper();
            

            foreach (var workbook in _appSettings.Projects.First(x => x.ProjectName == projectName).MetricsSheetPath)
            {
                var filePath = Path.Combine(_appSettings.OneDriveLocation, _appSettings.MetricsFolder, workbook);
                var recentFile = filePath.GetRecentlyModifiedSimilarFile();
                if (recentFile == null) continue;

                excelWrapper.Open(recentFile.FullName);
                var dashboardData = excelWrapper.ReadSpecificColumnsFromRange<DashboardData>("Dashboard", colVarPair, 3, "2:2").ToList();
                var match = dashboardData.FirstOrDefault(data => data.SprintName == sprintNameFormatted);
                if (match != null)
                    datas.Add(match);
            }

            List<SprintMetrics> metricsDatas = new();
            foreach (var data in datas)
            {
                var metrics = new SprintMetrics
                {
                    Committed = data.Assigned.ToLong(),
                    Delivered = data.Completed.ToLong(),
                    CommitmentIndex = data.Completed.ToLong() != 0 ? data.Assigned.ToLong() / data.Completed.ToLong() : 0,
                    Remarks = data.Remarks,
                    CodeReviewCommentsInternal = GetDefectsFromRemarks(data.Remarks, "External CRC"),
                    CodeReviewCommentsExternal = GetDefectsFromRemarks(data.Remarks, "Internal CRC"),
                    EscapedDefects = GetDefectsFromRemarks(data.Remarks, "Escaped Defects"),
                    QADefects = GetDefectsFromRemarks(data.Remarks, "QA Defects"),
                };
                metricsDatas.Add(metrics);
            }

            sprintMetrics = new SprintMetrics
            {
                Sprint = sprintName,
                LastSprint = "No",
                BacklogHealth = 30,
                Velocity = reportData.Reverse().Take(3).Average(x => x.Velocity.ToDouble()),
                CommitmentIndex = metricsDatas.Sum(x => x.Delivered.ToInt()) != 0
                    ? metricsDatas.Sum(x => x.Committed.ToInt()) / metricsDatas.Sum(x => x.Delivered.ToInt())
                    : 0,
                CodeReviewCommentsExternal = metricsDatas.Sum(x => x.CodeReviewCommentsExternal.ToInt()),
                CodeReviewCommentsInternal = metricsDatas.Sum(x => x.CodeReviewCommentsInternal.ToInt()),
                EscapedDefects = metricsDatas.Sum(x => x.EscapedDefects.ToInt()),
                QADefects = metricsDatas.Sum(x => x.QADefects.ToInt()),
                Committed = metricsDatas.Sum(x => x.Committed.ToInt()),
                Delivered = metricsDatas.Sum(x => x.Delivered.ToInt()),
                Remarks = reportData.FirstOrDefault()?.Remarks,
                CodeQualityIndex = $"{metricsDatas.Sum(x => x.CodeReviewCommentsExternal.ToInt()) + metricsDatas.Sum(x => x.CodeReviewCommentsInternal.ToInt())}%"
            };
        }*/

        // Always save chart image and pre-fill narrative fields, regardless of data source
        var path = Path.Combine(_appSettings.OneDriveLocation, _appSettings.ReportAndDataFolder,
            _appSettings.Projects.First(x => x.ProjectName == projectName).DataFileName);

        sprintMetrics.ImagePath = SaveChartImage(path, "Scorecard");

        // Check if a PPT report already exists and pre-fill narrative fields
        var presentation = new Presentation();
       

        var outputPPTPath = Path.Combine(
            Path.GetDirectoryName(Path.Combine(_appSettings.OneDriveLocation, _appSettings.ReportAndDataFolder,
                _appSettings.Projects.Single(x => x.ProjectName == projectName).DataFileName))!,
            $"GlobalPayments-{projectName}-DeliveryQualitySummaryReport-{sprintNameFormatted}.pptx");

        sprintMetrics.OutputPPTPath = outputPPTPath;

        if (File.Exists(outputPPTPath))
        {
            presentation.LoadFromFile(outputPPTPath);
            if (presentation.Slides.Count > 3)
                throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_300));

            var slide = presentation.Slides[1];

            if (slide.Shapes[6] is IAutoShape sprintSummaryShape)
                sprintMetrics.SprintSummary = sprintSummaryShape.TextFrame.Text.Replace("Sprint Delivery Summary", "").Split("\r");

            if (slide.Shapes[4] is IAutoShape sprintHighlights)
                sprintMetrics.SprintHighlights = sprintHighlights.TextFrame.Text.Replace("Highlights:", "").Split("\r");

            if (slide.Shapes[7] is IAutoShape sprintRetrospective)
                sprintMetrics.SprintRetrospective = sprintRetrospective.TextFrame.Text.Replace("Retrospective:", "").Split("\r");
        }
        else
        {
            //fetch data from the dashboard sheet on the metricssheetpath workbooks 
            var reportDaatTest = GetDashboardDataForRange(projectName, "B3:O90").ToList();

            var sprintData = reportDaatTest.FirstOrDefault(data => data[0]?.ToString() == sprintNameFormatted);
            //find the sprintname formaatted in reportdaatest
            //log error when sprint data is null
            if (sprintData == null)
            {
                sprintMetrics.SprintMetricsDataAvailable = false;
                 _logger.LogInformation($"Sprint data not found for sprint Name on Metrics sheet: {sprintNameFormatted}"); 
            }
         else
            {
                 sprintMetrics.SprintMetricsDataAvailable = true;
                
                {
                    sprintMetrics.SprintSummary = sprintData[6]?.ToString()?.Split("\n") ?? Array.Empty<string>();
                    sprintMetrics.SprintHighlights = sprintData[7]?.ToString()?.Split("\n") ?? Array.Empty<string>();
                   sprintMetrics.SprintRetrospective = sprintData[9]?.ToString()?.Split("\n") ?? Array.Empty<string>();
                    

                }
            }
        }

        return sprintMetrics;
    }

    private int GetDefectsFromRemarks(string? remarks, string extractString)
    {
        if (string.IsNullOrEmpty(remarks))
            return -1;

        int returnValue = -1;
        string[] lines = remarks.Split(["\r\n", "\r", "\n"], StringSplitOptions.RemoveEmptyEntries);

        foreach (string line in lines)
        {
            try
            {
                if (line.Contains(extractString, StringComparison.OrdinalIgnoreCase))
                {
                    string[] parts = line.Split('-');
                    if (parts.Length >= 2 && int.TryParse(parts[1].Trim(), out int value))
                        returnValue = value;
                }
            }
            catch
            {
                returnValue = -1;
            }
        }

        return returnValue;
    }

   

    public (bool bReturn, string pdfPath) GeneratePresentation(ReportDataParameters reportParams, bool generatePdf = true)
    {
        var presentation = new Presentation();
        var templatepath = _appSettings.SprintMetricsReportTemplatePath;        
        
        if (!File.Exists(templatepath))
            throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_103, templatepath));

        var project = _appSettings.Projects.SingleOrDefault(x => x.ProjectName == reportParams.ProjectName)
            ?? throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_400, reportParams.ProjectName));

        var outputPPTPath = Path.Combine(
            Path.GetDirectoryName(Path.Combine(_appSettings.OneDriveLocation, _appSettings.ReportAndDataFolder,
                project.DataFileName))!,
            $"GlobalPayments-{reportParams.ProjectName}-DeliveryQualitySummaryReport-{reportParams.SprintName}.pptx");
        var outputPDFPath = Path.Combine(_appSettings.TempPath,
            $"GlobalPayments-{reportParams.ProjectName}-DeliveryQualitySummaryReport-{reportParams.SprintName}.pdf");

        try
        {
            presentation.LoadFromFile(templatepath);
        }
        catch (Exception ex)
        {
            throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_301, $"Unable to load template '{templatepath}'. The file may be corrupt or locked. {ex.Message}"), ex);
        }
        if (presentation.Slides.Count > 3)
            throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_300));

        // Slide 1 - Title
        var slide1 = presentation.Slides[0];
        if (slide1.Shapes[0] is IAutoShape autoShape)
        {
            autoShape.TextFrame.Paragraphs.Clear();

            TextParagraph paragraph1 = new();
            TextRange textRange1 = new("Delivery Quality Summary Report");
            textRange1.FontHeight = 28;
            textRange1.IsBold = TriState.True;
            paragraph1.TextRanges.Append(textRange1);
            autoShape.TextFrame.Paragraphs.Append(paragraph1);

            TextParagraph paragraph2 = new();
            TextRange textRange2 = new($"{reportParams.ProjectName} - ");
            textRange2.FontHeight = 18;
            paragraph2.TextRanges.Append(textRange2);
            autoShape.TextFrame.Paragraphs.Append(paragraph2);

            TextParagraph paragraph3 = new();
            TextRange textRange3 = new($"{reportParams.SprintNameWithDate} - ");
            textRange3.FontHeight = 12;
            paragraph3.TextRanges.Append(textRange3);
            autoShape.TextFrame.Paragraphs.Append(paragraph3);
        }

        // Slide 2 - Content
        var slide2 = presentation.Slides[1];

        // Shape 3: Project name
        if (slide2.Shapes[3] is IAutoShape appShape)
            appShape.TextFrame.Text = reportParams.ProjectName;

        // Shape 6: Sprint Delivery Summary
        if (slide2.Shapes[6] is IAutoShape shapeSummary)
        {
            shapeSummary.TextFrame.Paragraphs.Clear();
            TextParagraph textParagraph = new();
            TextRange txtRange = new("Sprint Delivery Summary");
            txtRange.FontHeight = 12;
            txtRange.IsBold = TriState.True;
            textParagraph.TextRanges.Append(txtRange);
            shapeSummary.TextFrame.Paragraphs.Append(textParagraph);
            shapeSummary.TextFrame.Paragraphs.Append(new TextParagraph());

            if (reportParams.SprintSummary != null)
            {
                foreach (string item in reportParams.SprintSummary)
                {
                    TextParagraph txtpara = new();
                    TextRange range = new(item);
                    range.FontHeight = 10;
                    txtpara.BulletType = TextBulletType.Symbol;
                    txtpara.Indent = -15;
                    txtpara.TextRanges.Append(range);
                    shapeSummary.TextFrame.Paragraphs.Append(txtpara);
                }
            }
        }

        // Shape 4: Highlights
        if (slide2.Shapes[4] is IAutoShape shapeHighlights)
        {
            shapeHighlights.TextFrame.Paragraphs.Clear();
            TextParagraph textParagraph = new();
            TextRange txtRange = new("Highlights:");
            txtRange.FontHeight = 12;
            txtRange.IsBold = TriState.True;
            textParagraph.TextRanges.Append(txtRange);
            shapeHighlights.TextFrame.Paragraphs.Append(textParagraph);
            shapeHighlights.TextFrame.Paragraphs.Append(new TextParagraph());

            if (reportParams.SprintHighlights != null)
            {
                foreach (string item in reportParams.SprintHighlights)
                {
                    TextParagraph txtpara = new();
                    TextRange range = new(item);
                    range.FontHeight = 10;
                    txtpara.BulletType = TextBulletType.Symbol;
                    txtpara.Indent = -15;
                    txtpara.TextRanges.Append(range);
                    shapeHighlights.TextFrame.Paragraphs.Append(txtpara);
                }
            }
        }

        // Shape 7: Retrospective
        if (slide2.Shapes[7] is IAutoShape shapeRetrospective)
        {
            shapeRetrospective.TextFrame.Paragraphs.Clear();
            TextParagraph textParagraph = new();
            TextRange txtRange = new("Retrospective:");
            txtRange.FontHeight = 12;
            txtRange.IsBold = TriState.True;
            textParagraph.TextRanges.Append(txtRange);
            shapeRetrospective.TextFrame.Paragraphs.Append(textParagraph);
            shapeRetrospective.TextFrame.Paragraphs.Append(new TextParagraph());

            if (reportParams.SprintRetrospective != null)
            {
                List<string> wentWellItems = new();
                List<string> didntGoWellItems = new();
                List<string> improvementItems = new();
                string currentCategory = "";

                foreach (string item in reportParams.SprintRetrospective)
                {
                    if (item.StartsWith("what went well", StringComparison.OrdinalIgnoreCase))
                    { currentCategory = "WentWell"; continue; }
                    else if (item.StartsWith("what didn't go well", StringComparison.OrdinalIgnoreCase) ||
                             item.StartsWith("what didnt go well", StringComparison.OrdinalIgnoreCase))
                    { currentCategory = "DidntGoWell"; continue; }
                    else if (item.StartsWith("Improvements", StringComparison.OrdinalIgnoreCase))
                    { currentCategory = "Improvements"; continue; }

                    switch (currentCategory)
                    {
                        case "WentWell": wentWellItems.Add(item); break;
                        case "DidntGoWell": didntGoWellItems.Add(item); break;
                        case "Improvements": improvementItems.Add(item); break;
                    }
                }

                AddSectionWithBulletPoints(shapeRetrospective, "What went well?", wentWellItems);
                AddSectionWithBulletPoints(shapeRetrospective, "What didn't go well?", didntGoWellItems);
                AddSectionWithBulletPoints(shapeRetrospective, "Improvements:", improvementItems);
            }
        }

        // Shape 8: Score
        if (slide2.Shapes[8] is IAutoShape shapeScore)
            shapeScore.TextFrame.Text = reportParams.SprintScore;

        // Append chart image
        if (!File.Exists(reportParams.ImagePath))
            throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_303, reportParams.ImagePath));

        using var fileStream = new FileStream(reportParams.ImagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        IImageData img = presentation.Images.Append(fileStream);
        int width = 390;
        int height = 180;
        int x = 335;
        int y = 26;
        RectangleF rect = new(x, y, width, height);
        IEmbedImage imageShape = slide2.Shapes.AppendEmbedImage(ShapeType.Rectangle, img, rect);
        imageShape.Line.FillType = FillFormatType.None;

        try
        {
            presentation.SaveToFile(outputPPTPath, Spire.Presentation.FileFormat.Pptx2016);
        }
        catch (Exception ex)
        {
            throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_301, $"Unable to save PPTX to '{outputPPTPath}'. {ex.Message}"), ex);
        }

        if (generatePdf)
        {
            try
            {
                presentation.SaveToFile(outputPDFPath, Spire.Presentation.FileFormat.PDF);
            }
            catch (Exception ex)
            {
                throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_301, $"Unable to save PDF to '{outputPDFPath}'. {ex.Message}"), ex);
            }
        }

        return (true, generatePdf ? outputPDFPath : outputPPTPath);
    }

    private static void AddSectionWithBulletPoints(IAutoShape shape, string headerText, List<string> items)
    {
        if (items.Count == 0) return;

        TextParagraph headerParagraph = new();
        TextRange headerRange = new(headerText);
        headerRange.FontHeight = 11;
        headerRange.TextUnderlineType = TextUnderlineType.Single;
        headerParagraph.TextRanges.Append(headerRange);
        shape.TextFrame.Paragraphs.Append(headerParagraph);

        foreach (string item in items)
        {
            TextParagraph txtpara = new();
            TextRange range = new(ReplaceNumberingOnString(item));
            range.FontHeight = 10;
            txtpara.BulletType = TextBulletType.Symbol;
            txtpara.Indent = -15;
            txtpara.TextRanges.Append(range);
            shape.TextFrame.Paragraphs.Append(txtpara);
        }

        shape.TextFrame.Paragraphs.Append(new TextParagraph());
    }

    private static string ReplaceNumberingOnString(string inputText)
    {
        return string.IsNullOrEmpty(inputText) ? "" : Regex.Replace(inputText, @"^\d+\.\s", "", RegexOptions.Multiline);
    }

    private string SaveChartImage(string filePath, string sheetName)
    {
        var recentFile = filePath.GetRecentlyModifiedSimilarFile();
        if (recentFile == null) throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_302, filePath));

        var workbook = new Workbook();
        workbook.LoadFromFile(recentFile.FullName, ExcelVersion.Version2013);

        Spire.Xls.Chart? chart;
        try
        {
            chart = workbook.Worksheets[sheetName].Charts[0];
        }
        catch
        {
            throw new Exception(ErrorCodes.GetMessage(ErrorCodes.ERR_203));
        }

        chart.RefreshChart();
        var chartFileName = Path.Combine(_appSettings.TempPath, $"{chart.ChartTitle}.png");
        chart.SaveToImage(chartFileName);

        return chartFileName;
    }

    private IEnumerable<SprintMetrics> GetReportDataForProject(string projectName)
    {
        var dicColVar = new Dictionary<string, string>
        {
            { "Sprint", "Sprint" },
            { "Committed", "Committed" },
            { "Delivered", "Delivered" },
            { "Commitment Index", "CommitmentIndex" },
            { "Velocity", "Velocity" },
            { "Code Coverage", "CodeQualityIndex" },
            { "Code Review Comments - Internal", "CodeReviewCommentsInternal" },
            { "Code Review Comments - External", "CodeReviewCommentsExternal" },
            { "QA Defects", "QADefects" },
            { "Escaped Defects", "EscapedDefects" },
            { "Backlog Health", "BacklogHealth" },
            { "Last Sprint?", "LastSprint" },
            { "Remarks", "Remarks" }
        };

        using var wrapper = new ExcelWrapper();
        var filePath = Path.Combine(_appSettings.OneDriveLocation, _appSettings.ReportAndDataFolder,
            _appSettings.Projects.First(x => x.ProjectName == projectName).DataFileName);

        _logger.LogInformation("GetReportDataForProject - Path: {Path}", filePath);

        var recentFile = filePath.GetRecentlyModifiedSimilarFile();
        if (recentFile == null)
        {
            _logger.LogWarning("GetReportDataForProject - No file found for: {Path}", filePath);
            return [];
        }

        _logger.LogInformation("GetReportDataForProject - Opening: {File}", recentFile.FullName);
        wrapper.Open(recentFile.FullName);
        var data = wrapper.ReadSpecificColumnsFromRange<SprintMetrics>("Data", dicColVar, 1, "1:1").ToList();
        _logger.LogInformation("GetReportDataForProject - Rows read: {Count}", data.Count);
        return data;
    }

    private List<object?[]> GetDashboardDataForRange(string projectName, string range) 
    {
        var result = new List<object?[]>();
        var metricsPaths = _appSettings.Projects
            .First(x => x.ProjectName == projectName).MetricsSheetPath;

        foreach (var metricsPath in metricsPaths)
        {
            var filePath = Path.Combine(_appSettings.OneDriveLocation, _appSettings.MetricsFolder, metricsPath);
            _logger.LogInformation("GetDashboardDataForRange - Path: {Path}", filePath);

            var recentFile = filePath.GetRecentlyModifiedSimilarFile();
            if (recentFile == null)
            {
                _logger.LogWarning("GetProjectDataForRange - No file found for: {Path}", filePath);
                continue;
            }

            _logger.LogInformation("GetProjectDataForRange - Opening: {File}", recentFile.FullName);
            using var wrapper = new ExcelWrapper();
            wrapper.Open(recentFile.FullName);
            var data = wrapper.ReadRangeAsObjects("Dashboard", range);
            _logger.LogInformation("GetProjectDataForRange - Rows read from {File}: {Count}", recentFile.Name, data.Count);
            result.AddRange(data);
        }

        _logger.LogInformation("GetProjectDataForRange - Total rows after concatenation: {Count}", result.Count);
        return result;
    }
}
