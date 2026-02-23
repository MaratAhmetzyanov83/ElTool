// FILE: src/Integrations/ExcelGateway.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Define abstract gateway for Excel template I/O and provide default CSV-compatible fallback implementation.
//   SCOPE: INPUT sheet export and OUTPUT sheet import contract.
//   DEPENDS: M-CONFIG
//   LINKS: M-EXPORT, M-CONFIG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   IExcelGateway - Abstraction for excel read/write operations.
//   CsvExcelGateway - Fallback implementation writing/reading semicolon separated content.
// END_MODULE_MAP

using System.Globalization;
using System.Text;
using ElTools.Models;

namespace ElTools.Integrations;

public interface IExcelGateway
{
    void WriteInputRows(string templatePath, IReadOnlyList<ExcelInputRow> rows);
    IReadOnlyList<ExcelOutputRow> ReadOutputRows(string templatePath);
}

public sealed class CsvExcelGateway : IExcelGateway
{
    private const char Separator = ';';

    public void WriteInputRows(string templatePath, IReadOnlyList<ExcelInputRow> rows)
    {
        // START_BLOCK_WRITE_INPUT_ROWS
        string directory = Path.GetDirectoryName(templatePath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        string output = Path.Combine(string.IsNullOrWhiteSpace(directory) ? "." : directory, $"{fileName}.INPUT.csv");

        var lines = new List<string>
        {
            string.Join(Separator, new[]
            {
                "ЩИТ","ГРУППА","МОЩНОСТЬ_кВт","НАПРЯЖЕНИЕ","ДЛИНА_м","ДЛИНА_ПОТОЛОК_м","ДЛИНА_ПОЛ_м","ДЛИНА_СТОЯК_м","ФАЗА","ТИП_ГРУППЫ"
            })
        };

        foreach (ExcelInputRow row in rows)
        {
            lines.Add(string.Join(Separator, new[]
            {
                Escape(row.Shield),
                Escape(row.Group),
                row.PowerKw.ToString("0.###", CultureInfo.InvariantCulture),
                row.Voltage.ToString("0.###", CultureInfo.InvariantCulture),
                row.TotalLengthMeters.ToString("0.###", CultureInfo.InvariantCulture),
                row.CeilingLengthMeters.ToString("0.###", CultureInfo.InvariantCulture),
                row.FloorLengthMeters.ToString("0.###", CultureInfo.InvariantCulture),
                row.RiserLengthMeters.ToString("0.###", CultureInfo.InvariantCulture),
                Escape(row.Phase ?? string.Empty),
                Escape(row.GroupType ?? string.Empty)
            }));
        }

        File.WriteAllLines(output, lines, Encoding.UTF8);
        // END_BLOCK_WRITE_INPUT_ROWS
    }

    public IReadOnlyList<ExcelOutputRow> ReadOutputRows(string templatePath)
    {
        // START_BLOCK_READ_OUTPUT_ROWS
        string directory = Path.GetDirectoryName(templatePath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        string input = Path.Combine(string.IsNullOrWhiteSpace(directory) ? "." : directory, $"{fileName}.OUTPUT.csv");
        if (!File.Exists(input))
        {
            return [];
        }

        var rows = new List<ExcelOutputRow>();
        string[] lines = File.ReadAllLines(input, Encoding.UTF8);
        foreach (string line in lines.Skip(1))
        {
            string[] parts = line.Split(Separator);
            if (parts.Length < 7)
            {
                continue;
            }

            int breakerModules = TryParseInt(parts.ElementAtOrDefault(5));
            int rcdModules = TryParseInt(parts.ElementAtOrDefault(6));
            rows.Add(new ExcelOutputRow(
                parts[0],
                parts[1],
                parts[2],
                parts[3],
                parts[4],
                breakerModules,
                rcdModules,
                parts.ElementAtOrDefault(7)));
        }

        return rows;
        // END_BLOCK_READ_OUTPUT_ROWS
    }

    private static int TryParseInt(string? raw)
    {
        // START_BLOCK_TRY_PARSE_INT
        return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : 0;
        // END_BLOCK_TRY_PARSE_INT
    }

    private static string Escape(string value)
    {
        // START_BLOCK_ESCAPE_VALUE
        return value.Replace(Separator.ToString(), " ");
        // END_BLOCK_ESCAPE_VALUE
    }
}
