using ClosedXML.Excel;
using System.Text;

string path = @"C:\Users\marat\ElTool\tmp\range-import\Однолинейка ЩР ЭОМ.xlsx";
using var wb = new XLWorkbook(path);
var ws = wb.Worksheet("Диапазоны");
Console.OutputEncoding = Encoding.UTF8;
Console.WriteLine($"Sheet: {ws.Name}");
int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
Console.WriteLine($"Used: A1:{ToCol(lastCol)}{lastRow}");
for (int r = 1; r <= Math.Min(40,lastRow); r++)
{
    var parts = new List<string>();
    for (int c = 1; c <= Math.Min(8,lastCol); c++)
    {
        var cell = ws.Cell(r,c);
        string v = string.Empty;
        if (!cell.IsEmpty())
        {
            v = cell.GetFormattedString();
            if (string.IsNullOrWhiteSpace(v)) v = cell.GetString();
            v = v.Replace("\r"," ").Replace("\n"," ").Trim();
        }
        if (!string.IsNullOrWhiteSpace(v)) parts.Add($"{ToCol(c)}{r}='{v}'");
    }
    if (parts.Count > 0) Console.WriteLine(string.Join("; ", parts));
}

static string ToCol(int col)
{
    if (col <= 0) return "?";
    int x = col;
    var sb = new StringBuilder();
    while (x > 0)
    {
        x--; sb.Insert(0, (char)('A' + (x % 26))); x /= 26;
    }
    return sb.ToString();
}
