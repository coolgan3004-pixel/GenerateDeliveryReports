using System.Net;
using System.Text;
using GenerateDeliveryReports.Models;

namespace GenerateDeliveryReports.Worker;

public static class ReportEmailBuilder
{
    public static string BuildHtml(List<RunHistory> runHistory, DateTimeOffset cycleTime)
    {
        var sb = new StringBuilder();

        // Build HTML structure
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html lang='en'>");
        sb.AppendLine("<head>");
        sb.AppendLine("  <meta charset='UTF-8'>");
        sb.AppendLine("  <meta name='viewport' content='width=device-width, initial-scale=1.0'>");
        sb.AppendLine("  <title>Worker Run History Dashboard</title>");
        sb.AppendLine("  <style>");
        sb.AppendLine("    * { margin: 0; padding: 0; box-sizing: border-box; }");
        sb.AppendLine("    body {");
        sb.AppendLine("      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;");
        sb.AppendLine("      font-size: 14px;");
        sb.AppendLine("      color: #2c3e50;");
        sb.AppendLine("      background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);");
        sb.AppendLine("      min-height: 100vh;");
        sb.AppendLine("      padding: 20px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .container { max-width: 1200px; margin: 0 auto; }");
        sb.AppendLine("    .header {");
        sb.AppendLine("      background: white;");
        sb.AppendLine("      padding: 30px;");
        sb.AppendLine("      border-radius: 12px;");
        sb.AppendLine("      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07), 0 1px 3px rgba(0, 0, 0, 0.06);");
        sb.AppendLine("      margin-bottom: 30px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .header h1 {");
        sb.AppendLine("      font-size: 32px;");
        sb.AppendLine("      margin-bottom: 8px;");
        sb.AppendLine("      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);");
        sb.AppendLine("      -webkit-background-clip: text;");
        sb.AppendLine("      -webkit-text-fill-color: transparent;");
        sb.AppendLine("      background-clip: text;");
        sb.AppendLine("    }");
        sb.AppendLine("    .header-info { display: flex; gap: 30px; margin-top: 15px; font-size: 14px; color: #666; }");
        sb.AppendLine("    .header-info div { display: flex; align-items: center; gap: 8px; }");
        sb.AppendLine("    .header-info strong { color: #2c3e50; font-weight: 600; }");
        sb.AppendLine("    .projects-grid {");
        sb.AppendLine("      display: grid;");
        sb.AppendLine("      grid-template-columns: repeat(auto-fill, minmax(500px, 1fr));");
        sb.AppendLine("      gap: 20px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .project-card {");
        sb.AppendLine("      background: white;");
        sb.AppendLine("      border-radius: 12px;");
        sb.AppendLine("      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07), 0 1px 3px rgba(0, 0, 0, 0.06);");
        sb.AppendLine("      overflow: hidden;");
        sb.AppendLine("      transition: transform 0.2s, box-shadow 0.2s;");
        sb.AppendLine("    }");
        sb.AppendLine("    .project-card:hover {");
        sb.AppendLine("      transform: translateY(-4px);");
        sb.AppendLine("      box-shadow: 0 12px 16px rgba(0, 0, 0, 0.1), 0 4px 8px rgba(0, 0, 0, 0.08);");
        sb.AppendLine("    }");
        sb.AppendLine("    .project-header {");
        sb.AppendLine("      padding: 20px;");
        sb.AppendLine("      border-bottom: 2px solid #f0f0f0;");
        sb.AppendLine("      background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%);");
        sb.AppendLine("    }");
        sb.AppendLine("    .project-name {");
        sb.AppendLine("      font-size: 18px;");
        sb.AppendLine("      font-weight: 700;");
        sb.AppendLine("      margin-bottom: 12px;");
        sb.AppendLine("      color: #2c3e50;");
        sb.AppendLine("    }");
        sb.AppendLine("    .project-stats {");
        sb.AppendLine("      display: flex;");
        sb.AppendLine("      gap: 12px;");
        sb.AppendLine("      flex-wrap: wrap;");
        sb.AppendLine("    }");
        sb.AppendLine("    .stat-badge {");
        sb.AppendLine("      padding: 6px 12px;");
        sb.AppendLine("      border-radius: 6px;");
        sb.AppendLine("      font-size: 12px;");
        sb.AppendLine("      font-weight: 600;");
        sb.AppendLine("      display: inline-flex;");
        sb.AppendLine("      align-items: center;");
        sb.AppendLine("      gap: 4px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .badge-completed { background: #d1e7dd; color: #0f5132; }");
        sb.AppendLine("    .badge-errored { background: #f8d7da; color: #842029; }");
        sb.AppendLine("    .badge-missing { background: #fff3cd; color: #856404; }");
        sb.AppendLine("    .project-content {");
        sb.AppendLine("      max-height: 600px;");
        sb.AppendLine("      overflow-y: auto;");
        sb.AppendLine("    }");
        sb.AppendLine("    .run-item {");
        sb.AppendLine("      padding: 16px 20px;");
        sb.AppendLine("      border-top: 1px solid #f0f0f0;");
        sb.AppendLine("      transition: background 0.15s;");
        sb.AppendLine("    }");
        sb.AppendLine("    .run-item:hover { background: #fafafa; }");
        sb.AppendLine("    .run-timestamp {");
        sb.AppendLine("      font-size: 12px;");
        sb.AppendLine("      color: #999;");
        sb.AppendLine("      font-weight: 500;");
        sb.AppendLine("      margin-bottom: 8px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .sprints-list {");
        sb.AppendLine("      display: flex;");
        sb.AppendLine("      flex-direction: column;");
        sb.AppendLine("      gap: 6px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .sprint-result {");
        sb.AppendLine("      display: flex;");
        sb.AppendLine("      align-items: center;");
        sb.AppendLine("      gap: 8px;");
        sb.AppendLine("      font-size: 13px;");
        sb.AppendLine("      padding: 4px 0;");
        sb.AppendLine("    }");
        sb.AppendLine("    .sprint-status {");
        sb.AppendLine("      font-weight: 600;");
        sb.AppendLine("      min-width: 100px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .status-completed { color: #0f5132; }");
        sb.AppendLine("    .status-errored { color: #842029; }");
        sb.AppendLine("    .status-missing { color: #856404; }");
        sb.AppendLine("    .sprint-name {");
        sb.AppendLine("      flex: 1;");
        sb.AppendLine("      color: #2c3e50;");
        sb.AppendLine("      word-break: break-word;");
        sb.AppendLine("    }");
        sb.AppendLine("    .error-detail {");
        sb.AppendLine("      font-size: 12px;");
        sb.AppendLine("      color: #842029;");
        sb.AppendLine("      background: #fff5f7;");
        sb.AppendLine("      padding: 6px 8px;");
        sb.AppendLine("      border-left: 3px solid #dc3545;");
        sb.AppendLine("      margin-top: 4px;");
        sb.AppendLine("      border-radius: 2px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .no-data {");
        sb.AppendLine("      text-align: center;");
        sb.AppendLine("      padding: 40px 20px;");
        sb.AppendLine("      color: #999;");
        sb.AppendLine("      font-size: 16px;");
        sb.AppendLine("    }");
        sb.AppendLine("    .run-separator {");
        sb.AppendLine("      height: 1px;");
        sb.AppendLine("      background: linear-gradient(to right, transparent, #e0e0e0, transparent);");
        sb.AppendLine("      margin: 8px 0;");
        sb.AppendLine("    }");
        sb.AppendLine("    @media (max-width: 768px) {");
        sb.AppendLine("      .projects-grid { grid-template-columns: 1fr; }");
        sb.AppendLine("      .header h1 { font-size: 24px; }");
        sb.AppendLine("      .header-info { flex-direction: column; gap: 10px; }");
        sb.AppendLine("    }");
        sb.AppendLine("  </style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        sb.AppendLine("<div class='container'>");

        // Header
        sb.AppendLine("  <div class='header'>");
        sb.AppendLine("    <h1>📊 Worker Run History Dashboard</h1>");
        sb.AppendLine("    <div class='header-info'>");
        sb.AppendLine($"      <div>🕐 <strong>Last Updated:</strong> {cycleTime:yyyy-MM-dd HH:mm:ss}</div>");
        sb.AppendLine($"      <div>📈 <strong>Total Runs:</strong> {runHistory.Count}</div>");
        sb.AppendLine("    </div>");
        sb.AppendLine("  </div>");

        if (runHistory.Count == 0)
        {
            sb.AppendLine("  <div class='no-data'>");
            sb.AppendLine("    <p>No run history available yet. Worker runs will appear here.</p>");
            sb.AppendLine("  </div>");
        }
        else
        {
            // Get all unique projects
            var allProjects = runHistory
                .SelectMany(r => r.Results)
                .Select(r => r.ProjectName)
                .Distinct()
                .OrderBy(p => p)
                .ToList();

            if (allProjects.Count == 0)
            {
                sb.AppendLine("  <div class='no-data'>");
                sb.AppendLine("    <p>No projects found in history.</p>");
                sb.AppendLine("  </div>");
            }
            else
            {
                sb.AppendLine("  <div class='projects-grid'>");

                // For each project
                foreach (var projectName in allProjects)
                {
                    // Get all results for this project across all runs
                    var projectRunResults = runHistory
                        .Select(run => new
                        {
                            run.RunTime,
                            Results = run.Results.Where(r => r.ProjectName == projectName).ToList()
                        })
                        .Where(x => x.Results.Count > 0)
                        .OrderByDescending(x => x.RunTime)
                        .ToList();

                    if (projectRunResults.Count == 0) continue;

                    // Calculate summary stats for this project
                    var completedCount = projectRunResults.SelectMany(x => x.Results)
                        .Count(r => r.Outcome == SprintReportOutcome.Completed);
                    var erroredCount = projectRunResults.SelectMany(x => x.Results)
                        .Count(r => r.Outcome == SprintReportOutcome.Errored);
                    var missingCount = projectRunResults.SelectMany(x => x.Results)
                        .Count(r => r.Outcome == SprintReportOutcome.Missing);

                    // Project card
                    sb.AppendLine("    <div class='project-card'>");

                    // Header
                    sb.AppendLine("      <div class='project-header'>");
                    sb.AppendLine($"        <div class='project-name'>📁 {WebUtility.HtmlEncode(projectName)}</div>");
                    sb.AppendLine("        <div class='project-stats'>");
                    sb.AppendLine($"          <span class='stat-badge badge-completed'>✅ {completedCount}</span>");
                    sb.AppendLine($"          <span class='stat-badge badge-errored'>❌ {erroredCount}</span>");
                    sb.AppendLine($"          <span class='stat-badge badge-missing'>⚠️ {missingCount}</span>");
                    sb.AppendLine("        </div>");
                    sb.AppendLine("      </div>");

                    // Content
                    sb.AppendLine("      <div class='project-content'>");

                    // For each run of this project
                    foreach (var (index, runResult) in projectRunResults.Select((r, i) => (i, r)))
                    {
                        sb.AppendLine("        <div class='run-item'>");
                        sb.AppendLine($"          <div class='run-timestamp'>🕐 {runResult.RunTime:yyyy-MM-dd HH:mm:ss}</div>");
                        sb.AppendLine("          <div class='sprints-list'>");

                        // For each sprint result in this run
                        foreach (var result in runResult.Results)
                        {
                            var statusEmoji = result.Outcome switch
                            {
                                SprintReportOutcome.Completed => "✅",
                                SprintReportOutcome.Errored => "❌",
                                SprintReportOutcome.Missing => "⚠️",
                                _ => "❓"
                            };

                            var statusClass = result.Outcome switch
                            {
                                SprintReportOutcome.Completed => "status-completed",
                                SprintReportOutcome.Errored => "status-errored",
                                SprintReportOutcome.Missing => "status-missing",
                                _ => ""
                            };

                            sb.AppendLine("            <div class='sprint-result'>");
                            sb.AppendLine($"              <span class='sprint-status {statusClass}'>{statusEmoji} {result.Outcome}</span>");
                            sb.AppendLine($"              <span class='sprint-name'>{WebUtility.HtmlEncode(result.SprintName)}</span>");
                            sb.AppendLine("            </div>");

                            // Show error detail if there is one
                            if (result.Outcome == SprintReportOutcome.Errored && !string.IsNullOrWhiteSpace(result.Detail))
                            {
                                sb.AppendLine($"            <div class='error-detail'>⚠️ {WebUtility.HtmlEncode(result.Detail)}</div>");
                            }
                        }

                        sb.AppendLine("          </div>");
                        sb.AppendLine("        </div>");

                        // Add separator between runs (except last)
                        if (index < projectRunResults.Count - 1)
                        {
                            sb.AppendLine("        <div class='run-separator'></div>");
                        }
                    }

                    sb.AppendLine("      </div>");
                    sb.AppendLine("    </div>");
                }

                sb.AppendLine("  </div>");
            }
        }

        sb.AppendLine("</div>");
        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        return sb.ToString();
    }
}
