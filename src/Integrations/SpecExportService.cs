// FILE: src/Integrations/SpecExportService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Export specification dataset to AutoCAD table and CSV.
//   SCOPE: Table render in drawing and CSV file generation.
//   DEPENDS: M-SPEC, M-ACAD, M-LOG
//   LINKS: M-EXPORT, M-SPEC, M-ACAD, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   ToCsv - Writes specification rows to CSV.
//   ToAutoCadTable - Placeholder for AutoCAD table generation.
// END_MODULE_MAP

using ElTools.Models;
using ElTools.Services;

namespace ElTools.Integrations;

public class SpecExportService
{
    private readonly LogService _log = new();

    public string ToCsv(IReadOnlyList<SpecificationRow> rows, string? path = null)
    {
        // START_BLOCK_EXPORT_TO_CSV
        string output = path ?? $"spec_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
        var lines = new List<string> { "CableType,Group,TotalLength" };
        lines.AddRange(rows.Select(r => $"{r.CableType},{r.Group},{r.TotalLength:0.###}"));
        File.WriteAllLines(output, lines);
        _log.Write($"CSV сформирован: {output}");
        return output;
        // END_BLOCK_EXPORT_TO_CSV
    }

    public void ToAutoCadTable(IReadOnlyList<SpecificationRow> rows)
    {
        // START_BLOCK_EXPORT_TO_AUTOCAD_TABLE
        _ = rows;
        _log.Write("Экспорт в таблицу AutoCAD будет реализован в следующей итерации.");
        // END_BLOCK_EXPORT_TO_AUTOCAD_TABLE
    }
}
