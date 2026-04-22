using System.Net;
using System.Text;
using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Worker;

public static class ReportEmailBuilder
{
    public static string BuildHtml(IReadOnlyList<SprintReportResult> results, DateTimeOffset cycleTime)
    {
        var missing   = results.Where(r => r.Outcome == SprintReportOutcome.Missing).ToList();
        var completed = results.Where(r => r.Outcome == SprintReportOutcome.Completed).ToList();
        var errored   = results.Where(r => r.Outcome == SprintReportOutcome.Errored).ToList();

        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html><head><meta charset='utf-8'><style>");
        sb.AppendLine("  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; color: #333; max-width: 960px; margin: 24px auto; }");
        sb.AppendLine("  h2   { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 8px; }");
        sb.AppendLine("  h3   { margin-top: 32px; font-size: 15px; }");
        sb.AppendLine("  .summary { display: flex; gap: 14px; margin: 16px 0 24px; }");
        sb.AppendLine("  .badge   { padding: 5px 16px; border-radius: 20px; font-weight: 600; font-size: 13px; }");
        sb.AppendLine("  .missing   { background: #fff3cd; color: #856404; border: 1px solid #ffc107; }");
        sb.AppendLine("  .completed { background: #d1e7dd; color: #0f5132; border: 1px solid #198754; }");
        sb.AppendLine("  .errored   { background: #f8d7da; color: #842029; border: 1px solid #dc3545; }");
        sb.AppendLine("  table { border-collapse: collapse; width: 100%; margin-top: 10px; }");
        sb.AppendLine("  th    { background: #f0f0f0; text-align: left; padding: 8px 12px; border: 1px solid #ddd; }");
        sb.AppendLine("  td    { padding: 7px 12px; border: 1px solid #ddd; vertical-align: top; }");
        sb.AppendLine("  tr:nth-child(even) { background: #f9f9f9; }");
        sb.AppendLine("  .none { color: #888; font-style: italic; margin: 6px 0; }");
        sb.AppendLine("</style></head><body>");

        sb.AppendLine("<h2>Delivery Report Cycle Summary</h2>");
        sb.AppendLine($"<p>Cycle run at: <strong>{cycleTime:yyyy-MM-dd HH:mm:ss zzz}</strong></p>");

        sb.AppendLine("<div class='summary'>");
        sb.AppendLine($"  <span class='badge missing'>{missing.Count} Missing</span>");
        sb.AppendLine($"  <span class='badge completed'>{completed.Count} Completed</span>");
        sb.AppendLine($"  <span class='badge errored'>{errored.Count} Errored</span>");
        sb.AppendLine("</div>");

        AppendSection(sb, "Missing Reports", missing, "missing",
            headers: ["Project", "Sprint", "Filename", "Expected Path"],
            rowSelector: r => [r.ProjectName, r.SprintName, r.Detail is not null ? Path.GetFileName(r.Detail) : "—", r.Detail ?? "—"]);

        AppendSection(sb, "Completed Reports", completed, "completed",
            headers: ["Project", "Sprint", "Filename", "Full Path"],
            rowSelector: r => [r.ProjectName, r.SprintName, r.Detail is not null ? Path.GetFileName(r.Detail) : "—", r.Detail ?? "—"]);

        AppendSection(sb, "Errored Reports", errored, "errored",
            headers: ["Project", "Sprint", "Error", "Sprint Name (Metrics Sheet)", "Summary", "Highlights", "Retrospective"],
            rowSelector: r => [
                r.ProjectName,
                r.SprintName,
                r.Detail ?? "—",
                r.SprintMetricsSprintName ?? "—",
                r.SprintSummary is { Length: > 0 } ? string.Join(" | ", r.SprintSummary.Where(s => !string.IsNullOrWhiteSpace(s))) : "—",
                r.SprintHighlights is { Length: > 0 } ? string.Join(" | ", r.SprintHighlights.Where(s => !string.IsNullOrWhiteSpace(s))) : "—",
                r.SprintRetrospective is { Length: > 0 } ? string.Join(" | ", r.SprintRetrospective.Where(s => !string.IsNullOrWhiteSpace(s))) : "—"
            ]);

        sb.AppendLine("</body></html>");
        return sb.ToString();
    }

    private static void AppendSection(
        StringBuilder sb,
        string title,
        List<SprintReportResult> items,
        string cssClass,
        string[] headers,
        Func<SprintReportResult, string[]> rowSelector)
    {
        sb.AppendLine($"<h3><span class='badge {cssClass}'>{items.Count}</span>&nbsp; {WebUtility.HtmlEncode(title)}</h3>");

        if (items.Count == 0)
        {
            sb.AppendLine("<p class='none'>None</p>");
            return;
        }

        sb.AppendLine("<table><thead><tr>");
        foreach (var h in headers)
            sb.AppendLine($"  <th>{WebUtility.HtmlEncode(h)}</th>");
        sb.AppendLine("</tr></thead><tbody>");

        foreach (var item in items)
        {
            sb.AppendLine("  <tr>");
            foreach (var cell in rowSelector(item))
                sb.AppendLine($"    <td>{WebUtility.HtmlEncode(cell)}</td>");
            sb.AppendLine("  </tr>");
        }

        sb.AppendLine("</tbody></table>");
    }
}
