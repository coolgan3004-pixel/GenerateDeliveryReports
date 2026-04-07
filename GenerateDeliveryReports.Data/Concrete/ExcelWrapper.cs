using GenerateDeliveryReports.Data.Extensions;
using GenerateDeliveryReports.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Reflection;

namespace GenerateDeliveryReports.Data.Concrete;

public interface IWrapper : IDisposable
{
    void Open(string filePath);
    void Close();
    List<KeyValue> ReadRangeWithAddress(string sheetName, string range);
    List<T> ReadfromRangeAsCollection<T>(string sheetName, string range) where T : class;
    void WriteToRangeFromCollection<T>(string sheetName, string range, IEnumerable<T> collection) where T : class;
    object ReadCell(string sheetName, int row, int col);
    void UpdateCell(string sheetName, int row, int col, object value);
    void DeleteCell(string sheetName, int row, int col);
    IList<object[]> ReadRows(string sheetName, int startRow, int endRow);
    void WriteRow(string sheetName, int row, object[] values);
    void DeleteRow(string sheetName, int row);
    void Save();
}

public class ExcelWrapper : IWrapper
{
    private ExcelPackage? _package;
    private string _filePath = string.Empty;

    public void Open(string filePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("GenerateDeliveryReports");
        _filePath = filePath;
        var fileInfo = new FileInfo(filePath);
        _package = new ExcelPackage(fileInfo);
    }

    public void Close()
    {
        _package?.Dispose();
        _package = null;
    }

    public object ReadCell(string sheetName, int row, int col)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        return ws?.Cells[row, col].Value!;
    }

    public void Recalculate(string sheetName, string range)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        ws?.Cells[range].Calculate();
    }

    public void SetFormula(string sheetName, string range, string formula)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        ws!.Cells[range].Formula = formula;
        ws.Cells[range].Calculate();
    }

    public object ReadCellValue(string sheetName, string cellAddress)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        ws!.Cells[cellAddress].Calculate();
        return ws.Cells[cellAddress].Value;
    }

    public object ReadCellText(string sheetName, string cellAddress)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        ws!.Cells[cellAddress].Calculate();
        return ws.Cells[cellAddress].Text;
    }

    public void UpdateCell(string sheetName, int row, int col, object value)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        ws!.Cells[row, col].Value = value;
    }

    public void UpdateCell(string sheetName, string address, object value)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        ws!.Cells[address].Value = value;
    }

    public void DeleteCell(string sheetName, int row, int col)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        ws!.Cells[row, col].Clear();
    }

    public string GetAddressFromValue(string sheetName, string range, string value)
    {
        var values = ReadRangeWithAddress(sheetName, range);
        return values.Where(item => item.Key == value).SingleOrDefault()?.Value ?? string.Empty;
    }

    public List<T> ReadfromRangeAsCollection<T>(string sheetName, string range) where T : class
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        var cells = ws!.Cells[range];
        return cells.ToCollection<T>();
    }

    public List<object?[]> ReadRangeAsObjects(string sheetName, string range)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        var cells = ws!.Cells[range];
        var result = new List<object?[]>();

        for (int row = cells.Start.Row; row <= cells.End.Row; row++)
        {
            var rowValues = new object?[cells.End.Column - cells.Start.Column + 1];
            for (int col = cells.Start.Column; col <= cells.End.Column; col++)
                rowValues[col - cells.Start.Column] = ws.Cells[row, col].Value;
            result.Add(rowValues);
        }

        return result;
    }

    public IEnumerable<T> ReadSpecificColumnsFromRange<T>(string sheetName, Dictionary<string, string> dicColVariablePair, int rowStart = 1, string headerRange = "1:1") where T : class, new()
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        int rowCount = ws!.Dimension.Rows;
        List<T> values = new();

        for (int row = rowStart; row <= rowCount; row++)
        {
            var obj = new T();
            Type objType = typeof(T);
            var propId = objType.GetProperty("Id");
            propId?.SetValue(obj, row, null);
            int count = 0;

            foreach (var colVarPair in dicColVariablePair)
            {
                PropertyInfo? prop = objType.GetProperty(colVarPair.Value);
                var val = ws.Cells[row, ws.GetColumnByName(colVarPair.Key, headerRange)].Value;
                if (val == null)
                    count++;
                try
                {
                    prop?.SetValue(obj, val, null);
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower() == "object of type 'system.double' cannot be converted to type 'system.datetime'.")
                    {
                        DateTime dt = DateTime.Parse(DateTime.FromOADate(val!.ToLong()).ToString("MM-dd-yyyy"));
                        prop?.SetValue(obj, dt);
                    }
                }
            }
            if (count != dicColVariablePair.Count)
                values.Add(obj);
        }
        return values;
    }

    public void WriteToRangeFromCollection<T>(string sheetName, string range, IEnumerable<T> collection) where T : class
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        var cells = ws!.Cells[range];
        cells.Style.Numberformat.Format = "@";
        cells.LoadFromCollection(collection);
        cells.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        cells.AutoFitColumns();
    }

    public List<KeyValue> ReadRangeWithAddress(string sheetName, string range)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        var cells = ws!.Cells[range];
        return cells.Where(cell => cell?.Value != null)
            .Select(x => new KeyValue { Key = x.Value.ToString()!, Value = x.Address.ToString() })
            .ToList();
    }

    public IList<object[]> ReadRows(string sheetName, int startRow, int endRow)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        var result = new List<object[]>();
        for (int r = startRow; r <= endRow; r++)
        {
            int colCount = ws!.Dimension.End.Column;
            var row = new object[colCount];
            for (int c = 1; c <= colCount; c++)
                row[c - 1] = ws.Cells[r, c].Value;
            result.Add(row);
        }
        return result;
    }

    public void WriteRow(string sheetName, int row, object[] values)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        for (int c = 0; c < values.Length; c++)
            ws!.Cells[row, c + 1].Value = values[c];
    }

    public void DeleteRow(string sheetName, int row)
    {
        var ws = _package!.Workbook.Worksheets[sheetName];
        ws!.DeleteRow(row);
    }

    public void Save()
    {
        _package!.Save();
        _package.Dispose();
    }

    public void Dispose()
    {
        Close();
    }
}
