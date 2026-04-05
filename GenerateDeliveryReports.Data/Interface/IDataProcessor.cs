using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Data.Interface;

public interface IDataProcessor
{
    string GetEmailContent(string projectName, string sprintName);

    IEnumerable<string> GetSprintNames(string projectName);
    SprintMetrics? GetSprintMetrics(string projectName, string sprintName);
    IEnumerable<string> GetProjectNames();
    (bool bReturn, string pdfPath) GeneratePresentation(ReportDataParameters reportParams);
}
