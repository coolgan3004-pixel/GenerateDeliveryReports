namespace GenerateDeliveryReports.Models;

public class SprintInfo
{
    public string ProjectName { get; set; } = string.Empty;
    public string SprintName { get; set; } = string.Empty;
    public DateTime? SprintStartDate { get; set; }
    public DateTime? SprintEndDate { get; set; }
    public string OutputPPTPath { get; set; } = string.Empty;
}
