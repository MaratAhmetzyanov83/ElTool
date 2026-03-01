// FILE: src/Integrations/SpecExportService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Export specification dataset to AutoCAD table and CSV.
//   SCOPE: Table render in drawing and CSV file generation.
//   DEPENDS: M-AGGREGATION, M-CAD-CONTEXT, M-LOGGING, M-EXCEL-GATEWAY
//   LINKS: M-EXPORT, M-AGGREGATION, M-CAD-CONTEXT, M-LOGGING, M-EXCEL-GATEWAY
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
        Document? doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            var table = new Table
            {
                TableStyle = doc.Database.Tablestyle,
                Position = new Point3d(0, 0, 0)
            };
            table.SetSize(Math.Max(2, rows.Count + 1), 3);
            table.SetRowHeight(4);
            table.SetColumnWidth(28);

            table.Cells[0, 0].TextString = "Тип";
            table.Cells[0, 1].TextString = "Группа";
            table.Cells[0, 2].TextString = "Длина";

            for (int i = 0; i < rows.Count; i++)
            {
                SpecificationRow row = rows[i];
                table.Cells[i + 1, 0].TextString = row.CableType;
                table.Cells[i + 1, 1].TextString = row.Group;
                table.Cells[i + 1, 2].TextString = row.TotalLength.ToString("0.###");
            }

            ms.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            tr.Commit();
        }

        _log.Write($"Таблица AutoCAD сформирована. Строк: {rows.Count}.");
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

    public string ExportExcelOutputReportCsv(IReadOnlyList<ExcelOutputRow> rows, string? path = null)
    {
        // START_BLOCK_EXPORT_OUTPUT_REPORT_CSV
        string output = path ?? $"excel_output_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
        var lines = new List<string>
        {
            "ЩИТ,ГРУППА,АВТОМАТ,УЗО_ДИФ,КАБЕЛЬ,МОДУЛЕЙ_АВТОМАТ,МОДУЛЕЙ_УЗО,ПРИМЕЧАНИЕ"
        };
        lines.AddRange(rows.Select(r =>
            $"{r.Shield},{r.Group},{r.CircuitBreaker},{r.RcdDiff},{r.Cable},{r.CircuitBreakerModules},{r.RcdModules},{r.Note}"));
        File.WriteAllLines(output, lines);
        _log.Write($"CSV отчет по Excel OUTPUT сформирован: {output}");
        return output;
        // END_BLOCK_EXPORT_OUTPUT_REPORT_CSV
    }

    public void ToAutoCadTableFromOutput(IReadOnlyList<ExcelOutputRow> rows, Point3d position)
    {
        // START_BLOCK_EXPORT_OUTPUT_TO_AUTOCAD_TABLE
        Document? doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            var table = new Table
            {
                TableStyle = doc.Database.Tablestyle,
                Position = position
            };
            table.SetSize(Math.Max(2, rows.Count + 1), 6);
            table.SetRowHeight(4);
            table.SetColumnWidth(24);

            table.Cells[0, 0].TextString = "ЩИТ";
            table.Cells[0, 1].TextString = "ГРУППА";
            table.Cells[0, 2].TextString = "АВТОМАТ";
            table.Cells[0, 3].TextString = "УЗО/ДИФ";
            table.Cells[0, 4].TextString = "КАБЕЛЬ";
            table.Cells[0, 5].TextString = "МОДУЛИ";

            for (int i = 0; i < rows.Count; i++)
            {
                ExcelOutputRow row = rows[i];
                int moduleCount = Math.Max(1, row.CircuitBreakerModules + row.RcdModules);
                table.Cells[i + 1, 0].TextString = row.Shield;
                table.Cells[i + 1, 1].TextString = row.Group;
                table.Cells[i + 1, 2].TextString = row.CircuitBreaker;
                table.Cells[i + 1, 3].TextString = row.RcdDiff;
                table.Cells[i + 1, 4].TextString = row.Cable;
                table.Cells[i + 1, 5].TextString = moduleCount.ToString();
            }

            ms.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            tr.Commit();
        }

        _log.Write($"Таблица AutoCAD по Excel OUTPUT сформирована. Строк: {rows.Count}.");
        // END_BLOCK_EXPORT_OUTPUT_TO_AUTOCAD_TABLE
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

