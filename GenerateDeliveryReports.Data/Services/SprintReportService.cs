using GenerateDeliveryReports.Data.Interface;
using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Data.Services;

public class SprintReportService
{
    private readonly IDataProcessor _dataProcessor;

    public SprintReportService(IDataProcessor dataProcessor)
    {
        _dataProcessor = dataProcessor;
    }

    public IEnumerable<string> GetProjectNames() => _dataProcessor.GetProjectNames();

    public IEnumerable<string> GetSprintNames(string projectName) => _dataProcessor.GetSprintNames(projectName);

    public SprintMetrics? GetSprintMetrics(string projectName, string sprintName) => _dataProcessor.GetSprintMetrics(projectName, sprintName);


    public (bool bReturn, string pdfPath) GeneratePresentation(ReportDataParameters reportParams) => _dataProcessor.GeneratePresentation(reportParams);

    public string GetEmailContent(string projectName, string sprintName) => _dataProcessor.GetEmailContent(projectName, sprintName);
}
