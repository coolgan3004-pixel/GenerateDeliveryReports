using OfficeOpenXml;

namespace GenerateDeliveryReports.Data.Extensions;

public static class EPPlusExtensions
{
    public static int GetColumnByName(this ExcelWorksheet ws, string columnName, string headerRange)
    {
        if (ws == null) throw new ArgumentNullException(nameof(ws));
        var col = ws.Cells[headerRange].First(c => c.Value?.ToString() == columnName);
        return col?.Start.Column ?? -1;
    }
}
