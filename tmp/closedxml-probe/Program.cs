using ClosedXML.Excel;
using System.Globalization;

var path = @"C:\Users\marat\Desktop\Проект\04 Проект\№ 25-ЮН-ИНЖ 20ПИР-ЭОМ\Однолинейка ЩР ЭОМ.xlsx";
if (!File.Exists(path))
{
    Console.WriteLine($"FILE_NOT_FOUND: {path}");
    return;
}

using var wb = new XLWorkbook(path);
var ws = wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, "В Акад", StringComparison.OrdinalIgnoreCase))
         ?? wb.Worksheets.First();
Console.WriteLine($"SHEET={ws.Name}");

static string ReadCellText(IXLWorksheet sheet, int row, int column)
{
    var cell = sheet.Cell(row, column);
    if (cell.IsEmpty())
    {
        return string.Empty;
    }

    string value = cell.GetFormattedString();
    if (string.IsNullOrWhiteSpace(value))
    {
        value = cell.GetString();
    }

    return value.Replace('\u00A0', ' ').Trim();
}

static bool IsMeaningfulField(string value)
{
    if (string.IsNullOrWhiteSpace(value))
    {
        return false;
    }

    if (value.StartsWith('#') || value.StartsWith('='))
    {
        return false;
    }

    return !string.Equals(value, "Щит", StringComparison.OrdinalIgnoreCase)
        && !string.Equals(value, "Номер линии", StringComparison.OrdinalIgnoreCase)
        && !string.Equals(value, "Кабель", StringComparison.OrdinalIgnoreCase);
}

int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
Console.WriteLine($"LAST_ROW={lastRow}");

int accepted = 0;
for (int row = 2; row <= Math.Min(lastRow, 40); row++)
{
    string shield = ReadCellText(ws, row, 4);
    string group = ReadCellText(ws, row, 5);
    string note = ReadCellText(ws, row, 7);
    string breaker = ReadCellText(ws, row, 11);
    string cable = ReadCellText(ws, row, 13);

    bool ok = IsMeaningfulField(shield) && IsMeaningfulField(group) && IsMeaningfulField(cable)
              && !string.Equals(group, "0", StringComparison.OrdinalIgnoreCase);

    if (ok)
    {
        accepted++;
    }

    Console.WriteLine($"r{row}: D='{shield}' E='{group}' K='{breaker}' M='{cable}' | ok={ok}");
}

Console.WriteLine($"ACCEPTED_IN_PREVIEW={accepted}");

int totalAccepted = 0;
for (int row = 3; row <= lastRow; row++)
{
    string shield = ReadCellText(ws, row, 4);
    string group = ReadCellText(ws, row, 5);
    string cable = ReadCellText(ws, row, 13);

    if (!IsMeaningfulField(shield) || !IsMeaningfulField(group) || !IsMeaningfulField(cable))
    {
        continue;
    }

    if (string.Equals(group, "0", StringComparison.OrdinalIgnoreCase))
    {
        continue;
    }

    totalAccepted++;
}

Console.WriteLine($"TOTAL_ACCEPTED={totalAccepted}");
