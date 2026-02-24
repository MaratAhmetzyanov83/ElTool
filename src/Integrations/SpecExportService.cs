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
    private static readonly Dictionary<string, CachedExcelOutput> OutputCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly LogService _log = new();
    private readonly IExcelGateway _excelGateway = new CsvExcelGateway();

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

    public string ToExcelInput(string templatePath, IReadOnlyList<ExcelInputRow> rows)
    {
        // START_BLOCK_EXPORT_TO_EXCEL_INPUT
        _excelGateway.WriteInputRows(templatePath, rows);
        _log.Write($"Excel INPUT экспортирован на основе шаблона: {templatePath}");
        return templatePath;
        // END_BLOCK_EXPORT_TO_EXCEL_INPUT
    }

    public IReadOnlyList<ExcelOutputRow> FromExcelOutput(string templatePath)
    {
        // START_BLOCK_IMPORT_FROM_EXCEL_OUTPUT
        IReadOnlyList<ExcelOutputRow> rows = _excelGateway.ReadOutputRows(templatePath);
        OutputCache[NormalizePath(templatePath)] = new CachedExcelOutput(DateTime.UtcNow, rows.ToList());
        _log.Write($"Excel OUTPUT импортирован. Строк: {rows.Count}.");
        return rows;
        // END_BLOCK_IMPORT_FROM_EXCEL_OUTPUT
    }

    public IReadOnlyList<ExcelOutputRow> GetCachedOrLoadOutput(string templatePath)
    {
        // START_BLOCK_GET_CACHED_OR_LOAD_OUTPUT
        string key = NormalizePath(templatePath);
        if (OutputCache.TryGetValue(key, out CachedExcelOutput? cached))
        {
            _log.Write($"Использован кэш Excel OUTPUT. Строк: {cached.Rows.Count}.");
            return cached.Rows;
        }

        return FromExcelOutput(templatePath);
        // END_BLOCK_GET_CACHED_OR_LOAD_OUTPUT
    }

    public bool TryGetCachedOutput(string templatePath, out IReadOnlyList<ExcelOutputRow> rows)
    {
        // START_BLOCK_TRY_GET_CACHED_OUTPUT
        string key = NormalizePath(templatePath);
        if (OutputCache.TryGetValue(key, out CachedExcelOutput? cached))
        {
            rows = cached.Rows;
            return true;
        }

        rows = Array.Empty<ExcelOutputRow>();
        return false;
        // END_BLOCK_TRY_GET_CACHED_OUTPUT
    }

    public void ClearCachedOutput(string templatePath)
    {
        // START_BLOCK_CLEAR_CACHED_OUTPUT
        _ = OutputCache.Remove(NormalizePath(templatePath));
        // END_BLOCK_CLEAR_CACHED_OUTPUT
    }

    private static string NormalizePath(string path)
    {
        // START_BLOCK_NORMALIZE_PATH
        return Path.GetFullPath(path);
        // END_BLOCK_NORMALIZE_PATH
    }

    private sealed record CachedExcelOutput(DateTime ImportedAtUtc, IReadOnlyList<ExcelOutputRow> Rows);
}
