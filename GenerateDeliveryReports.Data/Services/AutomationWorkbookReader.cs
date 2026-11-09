using System.Globalization;
using GenerateDeliveryReports.Models;
using OfficeOpenXml;

namespace GenerateDeliveryReports.Data.Services;

public interface IAutomationWorkbookReader
{
    Task<IReadOnlyList<string>> GetWeeklySheetsAsync(Stream workbook, CancellationToken cancellationToken = default);
    Task<AutomationWorkbookReport> ReadAsync(Stream workbook, AutomationReportSelection selection, CancellationToken cancellationToken = default);
}

public sealed class AutomationWorkbookReader : IAutomationWorkbookReader
{
    private const int FirstProjectRow = 14;
    private const int LastProjectRow = 17;

    public async Task<IReadOnlyList<string>> GetWeeklySheetsAsync(Stream workbook, CancellationToken cancellationToken = default)
    {
        ExcelPackage.License.SetNonCommercialPersonal("GenerateDeliveryReports");
        using var package = new ExcelPackage();
        await package.LoadAsync(workbook, cancellationToken);
        return package.Workbook.Worksheets.Select(sheet => new { sheet.Name, Date = TryGetSheetDate(sheet.Name) })
            .Where(item => item.Date.HasValue).OrderByDescending(item => item.Date).Select(item => item.Name).ToList();
    }

    public async Task<AutomationWorkbookReport> ReadAsync(Stream workbook, AutomationReportSelection selection, CancellationToken cancellationToken = default)
    {
        ExcelPackage.License.SetNonCommercialPersonal("GenerateDeliveryReports");
        using var package = new ExcelPackage();
        await package.LoadAsync(workbook, cancellationToken);

        var datedSheets = package.Workbook.Worksheets
            .Select(sheet => new { Sheet = sheet, Date = TryGetSheetDate(sheet.Name) })
            .Where(item => item.Date.HasValue)
            .OrderBy(item => item.Date)
            .ToList();
        var current = datedSheets.SingleOrDefault(item => string.Equals(item.Sheet.Name, selection.CurrentSheetName, StringComparison.Ordinal));
        if (current is null || current.Date is null)
            throw new InvalidOperationException("Select a valid weekly input sheet.");

        var previous = datedSheets.SingleOrDefault(item => string.Equals(item.Sheet.Name, selection.PreviousSheetName, StringComparison.Ordinal));
        if (previous is null || previous.Date is null)
            throw new InvalidOperationException("Select a valid previous weekly input sheet.");
        if (previous.Date >= current.Date)
            throw new InvalidOperationException("The current week must be later than the comparison week.");
        var warnings = new List<string>();
        var headerDate = ReadDate(current.Sheet.Cells[6, 2].Value);
        if (headerDate.HasValue && DateOnly.FromDateTime(headerDate.Value) != current.Date.Value)
            warnings.Add($"Cell B6 ({headerDate:dd MMM yyyy}) differs from the selected sheet date ({current.Date.Value:dd MMM yyyy}). The sheet date is used for this report.");

        var projects = ReadProjects(current.Sheet)
            .Select(metric => new AutomationProjectReport
            {
                Current = metric,
                Previous = previous is null ? null : ReadProjects(previous.Sheet).FirstOrDefault(p => Normalize(p.Project) == Normalize(metric.Project))
            })
            .ToList();

        var report = new AutomationWorkbookReport
        {
            Client = ReadText(current.Sheet.Cells[5, 2].Value) ?? string.Empty,
            Team = ReadText(current.Sheet.Cells[7, 2].Value) ?? string.Empty,
            CurrentSheetName = current.Sheet.Name,
            CurrentWeek = current.Date.Value,
            PreviousSheetName = previous?.Sheet.Name,
            PreviousWeek = previous?.Date,
            Projects = projects,
            Warnings = warnings
        };
        return AutomationMetricValidator.Validate(report);
    }

    private static List<AutomationProjectMetrics> ReadProjects(ExcelWorksheet sheet) => Enumerable.Range(FirstProjectRow, LastProjectRow - FirstProjectRow + 1)
        .Select(row => ReadProject(sheet, row))
        .Where(project => !string.IsNullOrWhiteSpace(project.Project))
        .ToList();

    private static AutomationProjectMetrics ReadProject(ExcelWorksheet sheet, int row)
    {
        var coverage = sheet.Cells[row, 17];
        var passRate = sheet.Cells[row, 18];
        return new AutomationProjectMetrics
        {
            SourceRow = row,
            Project = ReadText(sheet.Cells[row, 1].Value) ?? string.Empty,
            Sprint = ReadText(sheet.Cells[row, 2].Value), Status = ReadText(sheet.Cells[row, 3].Value),
            ManualTests = ReadDecimal(sheet.Cells[row, 4].Value), CanBeAutomated = ReadDecimal(sheet.Cells[row, 5].Value), Automated = ReadDecimal(sheet.Cells[row, 6].Value), AddedThisSprint = ReadDecimal(sheet.Cells[row, 7].Value),
            TargetMinimum = ReadDecimal(sheet.Cells[row, 8].Value), TargetMaximum = ReadDecimal(sheet.Cells[row, 9].Value),
            Passed = ReadDecimal(sheet.Cells[row, 10].Value), TimedOut = ReadDecimal(sheet.Cells[row, 11].Value), Failed = ReadDecimal(sheet.Cells[row, 12].Value),
            LowDefects = ReadDecimal(sheet.Cells[row, 13].Value), MediumDefects = ReadDecimal(sheet.Cells[row, 14].Value), HighDefects = ReadDecimal(sheet.Cells[row, 15].Value),
            Comment = ReadText(sheet.Cells[row, 16].Value), EnteredCoverage = ReadDecimal(coverage.Value), EnteredPassRate = ReadDecimal(passRate.Value),
            CalculatedCoverage = CalculateRatio(ReadDecimal(sheet.Cells[row, 6].Value), ReadDecimal(sheet.Cells[row, 5].Value)),
            CalculatedPassRate = CalculateRatio(ReadDecimal(sheet.Cells[row, 10].Value), Sum(ReadDecimal(sheet.Cells[row, 10].Value), ReadDecimal(sheet.Cells[row, 11].Value), ReadDecimal(sheet.Cells[row, 12].Value)))
        };
    }

    private static decimal? Sum(params decimal?[] values) => values.Any(v => !v.HasValue) ? null : values.Sum(v => v!.Value);
    private static decimal? CalculateRatio(decimal? numerator, decimal? denominator) => numerator.HasValue && denominator is > 0 ? numerator.Value / denominator.Value : null;
    private static string? ReadText(object? value) => value?.ToString()?.Trim();
    private static decimal? ReadDecimal(object? value) => value is null ? null : decimal.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var result) ? result : null;
    private static DateTime? ReadDate(object? value) => value is DateTime date ? date : DateTime.TryParse(value?.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed) ? parsed : null;
    private static string Normalize(string value) => string.Concat(value.Where(c => !char.IsWhiteSpace(c))).ToUpperInvariant();
    private static DateOnly? TryGetSheetDate(string name)
    {
        var tail = name.Replace("Weekly Input -", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
        foreach (var format in new[] { "MMM dd yyyy", "MMM dd, yyyy", "MMM dd" })
        {
            if (!DateTime.TryParseExact(tail, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) continue;
            return DateOnly.FromDateTime(date.Year == 1 ? date.AddYears(DateTime.Today.Year - 1) : date);
        }
        return null;
    }
}

public static class AutomationMetricValidator
{
    public static AutomationWorkbookReport Validate(AutomationWorkbookReport report)
    {
        var validations = new List<CalculatedMetricValidation>();
        var projects = report.Projects.Select(project =>
        {
            var current = project.Current;
            validations.AddRange(ValidateMetric(current, "Q", "Coverage", current.EnteredCoverage, current.CalculatedCoverage));
            validations.AddRange(ValidateMetric(current, "R", "Pass rate", current.EnteredPassRate, current.CalculatedPassRate));
            var comparison = project.Previous is null ? null : new AutomationWeeklyComparison(
                Difference(current.CalculatedCoverage, project.Previous.CalculatedCoverage), Difference(current.CalculatedPassRate, project.Previous.CalculatedPassRate),
                Difference(current.Automated, project.Previous.Automated), Difference(current.AddedThisSprint, project.Previous.AddedThisSprint));
            return new AutomationProjectReport { Current = current, Previous = project.Previous, Comparison = comparison };
        }).ToList();

        return new AutomationWorkbookReport
        {
            Client = report.Client, Team = report.Team, CurrentSheetName = report.CurrentSheetName, CurrentWeek = report.CurrentWeek,
            PreviousSheetName = report.PreviousSheetName, PreviousWeek = report.PreviousWeek, Projects = projects,
            Warnings = report.Warnings, Validations = validations
        };
    }

    private static IEnumerable<CalculatedMetricValidation> ValidateMetric(AutomationProjectMetrics metrics, string column, string name, decimal? entered, decimal? calculated)
    {
        if (!entered.HasValue || !calculated.HasValue)
        {
            yield return new CalculatedMetricValidation(metrics.Project, $"{column}{metrics.SourceRow}", name, entered, calculated, "Workbook value", ValidationSeverity.Warning, !entered.HasValue ? "The workbook has no saved percentage value." : "The percentage cannot be calculated from the available counts.");
            yield break;
        }
        var mismatch = Math.Abs(entered.Value - calculated.Value) > 0.0005m;
        if (mismatch)
            yield return new CalculatedMetricValidation(metrics.Project, $"{column}{metrics.SourceRow}", name, entered, calculated, "Workbook value", ValidationSeverity.Conflict, $"Entered {name.ToLowerInvariant()} does not match the independently calculated value.");
    }

    private static decimal? Difference(decimal? current, decimal? previous) => current.HasValue && previous.HasValue ? current.Value - previous.Value : null;
}

