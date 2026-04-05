namespace GenerateDeliveryReports.Models;

public class DashboardData
{
    public int Id { get; set; }
    public string SprintName { get; set; } = string.Empty;
    public string? SprintSummary { get; set; }
    public string? SprintAccomplishment { get; set; }
    public string? SprintRetrospective { get; set; }
    public DateTime SprintStartDate { get; set; }
    public DateTime SprintEndDate { get; set; }
    public object? Assigned { get; set; }
    public object? Completed { get; set; }
    public string? Remarks { get; set; }
}
