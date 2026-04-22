using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Data.Interface;

public interface IDataProcessor
{
    string GetEmailContent(string projectName, string sprintName);

    IEnumerable<string> GetSprintNames(string projectName);
    IEnumerable<SprintInfo> GetSprintNamesWithDate(string projectName);
    SprintMetrics? GetSprintMetrics(SprintInfo sprint);
    IEnumerable<string> GetProjectNames();
    (bool bReturn, string pdfPath) GeneratePresentation(ReportDataParameters reportParams, bool generatePdf = true);
}
