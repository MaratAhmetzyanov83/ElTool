// FILE: src/Integrations/SpecExportService.cs
// VERSION: 1.2.1
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
//   TryInsertExcelLinkedTable - Inserts AutoCAD table linked to Excel range for visual parity.
// END_MODULE_MAP
//
// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.2.1 - Added runtime-safe named-range discovery for DataLink token resolution to support mixed ClosedXML API versions.
// END_CHANGE_SUMMARY

using System.Collections;
using System.Reflection;
using ClosedXML.Excel;
using ElTools.Models;
using ElTools.Services;

namespace ElTools.Integrations;

public class SpecExportService
{
    private const string ExcelDataAdapterId = "AcExcel";
    private const string ExcelAcadSheetName = "Р вЂ™ Р С’Р С”Р В°Р Т‘";
    private const string ExcelPrimaryRangeName = "Р В©Р С‘РЎвЂљ_1";
    private const string ExcelFallbackRangeName = "Р вЂќР В»РЎРЏ_Р Т‘Р С‘Р В°Р С—Р В°Р В·Р С•Р Р…Р С•Р Р†";
    private const string ExcelFallbackRangeAddress = "B3:M4000";
    // Captured from working DWG links: keeps Excel-driven width/height updates and source sync.
    private const int ExcelLinkUpdateOptionsMask = 68943873;

    private static readonly Dictionary<string, CachedExcelOutput> OutputCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly LogService _log = new();
    private readonly IExcelGateway _excelGateway = new CsvExcelGateway();

    // START_CONTRACT: ToCsv
    //   PURPOSE: To csv.
    //   INPUTS: { rows: IReadOnlyList<SpecificationRow> - method parameter; path: string? - method parameter }
    //   OUTPUTS: { string - textual result for to csv }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: ToCsv

    public string ToCsv(IReadOnlyList<SpecificationRow> rows, string? path = null)
    {
        // START_BLOCK_EXPORT_TO_CSV
        string output = path ?? $"spec_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
        var lines = new List<string> { "CableType,Group,TotalLength" };
        lines.AddRange(rows.Select(r => $"{r.CableType},{r.Group},{r.TotalLength:0.###}"));
        File.WriteAllLines(output, lines);
        _log.Write($"CSV РЎРѓРЎвЂћР С•РЎР‚Р СР С‘РЎР‚Р С•Р Р†Р В°Р Р…: {output}");
        return output;
        // END_BLOCK_EXPORT_TO_CSV
    }

    // START_CONTRACT: ToAutoCadTable
    //   PURPOSE: To auto cad table.
    //   INPUTS: { rows: IReadOnlyList<SpecificationRow> - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: ToAutoCadTable

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

            table.Cells[0, 0].TextString = "Р СћР С‘Р С—";
            table.Cells[0, 1].TextString = "Р вЂњРЎР‚РЎС“Р С—Р С—Р В°";
            table.Cells[0, 2].TextString = "Р вЂќР В»Р С‘Р Р…Р В°";

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

        _log.Write($"Р СћР В°Р В±Р В»Р С‘РЎвЂ Р В° AutoCAD РЎРѓРЎвЂћР С•РЎР‚Р СР С‘РЎР‚Р С•Р Р†Р В°Р Р…Р В°. Р РЋРЎвЂљРЎР‚Р С•Р С”: {rows.Count}.");
        // END_BLOCK_EXPORT_TO_AUTOCAD_TABLE
    }

    // START_CONTRACT: ToExcelInput
    //   PURPOSE: To excel input.
    //   INPUTS: { templatePath: string - method parameter; rows: IReadOnlyList<ExcelInputRow> - method parameter }
    //   OUTPUTS: { string - textual result for to excel input }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: ToExcelInput

    public string ToExcelInput(string templatePath, IReadOnlyList<ExcelInputRow> rows)
    {
        // START_BLOCK_EXPORT_TO_EXCEL_INPUT
        _excelGateway.WriteInputRows(templatePath, rows);
        _log.Write($"Excel INPUT РЎРЊР С”РЎРѓР С—Р С•РЎР‚РЎвЂљР С‘РЎР‚Р С•Р Р†Р В°Р Р… Р Р…Р В° Р С•РЎРѓР Р…Р С•Р Р†Р Вµ РЎв‚¬Р В°Р В±Р В»Р С•Р Р…Р В°: {templatePath}");
        return templatePath;
        // END_BLOCK_EXPORT_TO_EXCEL_INPUT
    }

    // START_CONTRACT: FromExcelOutput
    //   PURPOSE: From excel output.
    //   INPUTS: { templatePath: string - method parameter }
    //   OUTPUTS: { IReadOnlyList<ExcelOutputRow> - result of from excel output }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: FromExcelOutput

    public IReadOnlyList<ExcelOutputRow> FromExcelOutput(string templatePath)
    {
        // START_BLOCK_IMPORT_FROM_EXCEL_OUTPUT
        IReadOnlyList<ExcelOutputRow> rows = _excelGateway.ReadOutputRows(templatePath);
        OutputCache[NormalizePath(templatePath)] = new CachedExcelOutput(DateTime.UtcNow, rows.ToList());
        _log.Write($"Excel OUTPUT Р С‘Р СР С—Р С•РЎР‚РЎвЂљР С‘РЎР‚Р С•Р Р†Р В°Р Р…. Р РЋРЎвЂљРЎР‚Р С•Р С”: {rows.Count}.");
        if (rows.Count == 0 && _excelGateway is CsvExcelGateway csvGateway)
        {
            CsvExcelGateway.WorkbookReadDiagnostics? diagnostics = csvGateway.GetLastWorkbookReadDiagnostics();
            if (diagnostics is not null)
            {
                LogWorkbookReadDiagnostics(diagnostics);
            }
        }

        return rows;
        // END_BLOCK_IMPORT_FROM_EXCEL_OUTPUT
    }

    // START_CONTRACT: ExportExcelOutputReportCsv
    //   PURPOSE: Export excel output report csv.
    //   INPUTS: { rows: IReadOnlyList<ExcelOutputRow> - method parameter; path: string? - method parameter }
    //   OUTPUTS: { string - textual result for export excel output report csv }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: ExportExcelOutputReportCsv

    public string ExportExcelOutputReportCsv(IReadOnlyList<ExcelOutputRow> rows, string? path = null)
    {
        // START_BLOCK_EXPORT_OUTPUT_REPORT_CSV
        string output = path ?? $"excel_output_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
        var lines = new List<string>
        {
            "Р В©Р ВР Сћ,Р вЂњР В Р Р€Р СџР СџР С’,Р С’Р вЂ™Р СћР С›Р СљР С’Р Сћ,Р Р€Р вЂ”Р С›_Р вЂќР ВР В¤,Р С™Р С’Р вЂР вЂўР вЂєР В¬,Р СљР С›Р вЂќР Р€Р вЂєР вЂўР в„ў_Р С’Р вЂ™Р СћР С›Р СљР С’Р Сћ,Р СљР С›Р вЂќР Р€Р вЂєР вЂўР в„ў_Р Р€Р вЂ”Р С›,Р СџР В Р ВР СљР вЂўР В§Р С’Р СњР ВР вЂў"
        };
        lines.AddRange(rows.Select(r =>
            $"{r.Shield},{r.Group},{r.CircuitBreaker},{r.RcdDiff},{r.Cable},{r.CircuitBreakerModules},{r.RcdModules},{r.Note}"));
        File.WriteAllLines(output, lines);
        _log.Write($"CSV Р С•РЎвЂљРЎвЂЎР ВµРЎвЂљ Р С—Р С• Excel OUTPUT РЎРѓРЎвЂћР С•РЎР‚Р СР С‘РЎР‚Р С•Р Р†Р В°Р Р…: {output}");
        return output;
        // END_BLOCK_EXPORT_OUTPUT_REPORT_CSV
    }

    // START_CONTRACT: ToAutoCadTableFromOutput
    //   PURPOSE: To auto cad table from output.
    //   INPUTS: { rows: IReadOnlyList<ExcelOutputRow> - method parameter; position: Point3d - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: ToAutoCadTableFromOutput

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

            table.Cells[0, 0].TextString = "Р В©Р ВР Сћ";
            table.Cells[0, 1].TextString = "Р вЂњР В Р Р€Р СџР СџР С’";
            table.Cells[0, 2].TextString = "Р С’Р вЂ™Р СћР С›Р СљР С’Р Сћ";
            table.Cells[0, 3].TextString = "Р Р€Р вЂ”Р С›/Р вЂќР ВР В¤";
            table.Cells[0, 4].TextString = "Р С™Р С’Р вЂР вЂўР вЂєР В¬";
            table.Cells[0, 5].TextString = "Р СљР С›Р вЂќР Р€Р вЂєР В";

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

        _log.Write($"Р СћР В°Р В±Р В»Р С‘РЎвЂ Р В° AutoCAD Р С—Р С• Excel OUTPUT РЎРѓРЎвЂћР С•РЎР‚Р СР С‘РЎР‚Р С•Р Р†Р В°Р Р…Р В°. Р РЋРЎвЂљРЎР‚Р С•Р С”: {rows.Count}.");
        // END_BLOCK_EXPORT_OUTPUT_TO_AUTOCAD_TABLE
    }

    // START_CONTRACT: TryInsertExcelLinkedTable
    //   PURPOSE: Attempt to execute insert excel linked table.
    //   INPUTS: { workbookPath: string - method parameter; position: Point3d - method parameter; statusMessage: out string - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute insert excel linked table }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: TryInsertExcelLinkedTable

    public bool TryInsertExcelLinkedTable(string workbookPath, Point3d position, out string statusMessage)
    {
        // START_BLOCK_TRY_INSERT_EXCEL_LINKED_TABLE
        statusMessage = string.Empty;
        if (string.IsNullOrWhiteSpace(workbookPath))
        {
            statusMessage = "Р СџРЎС“РЎвЂљРЎРЉ Р С” Excel Р Р…Р Вµ Р В·Р В°Р Т‘Р В°Р Р….";
            return false;
        }

        string normalizedWorkbookPath = NormalizePath(workbookPath);
        if (!File.Exists(normalizedWorkbookPath))
        {
            statusMessage = $"Р В¤Р В°Р в„–Р В» Excel Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р…: {normalizedWorkbookPath}";
            return false;
        }

        Document? doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            statusMessage = "Р СњР ВµРЎвЂљ Р В°Р С”РЎвЂљР С‘Р Р†Р Р…Р С•Р С–Р С• Р Т‘Р С•Р С”РЎС“Р СР ВµР Р…РЎвЂљР В° AutoCAD.";
            return false;
        }

        string rangeToken = ResolvePreferredExcelRangeToken(normalizedWorkbookPath);
        string connectionPath = BuildWorkbookConnectionPath(doc.Name, normalizedWorkbookPath);
        string connectionString = $"{connectionPath}!{ExcelAcadSheetName}!{rangeToken}";
        string dataLinkName = BuildExcelDataLinkName(normalizedWorkbookPath, rangeToken);

        try
        {
            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                ObjectId linkId = EnsureExcelDataLink(tr, doc.Database, dataLinkName, connectionString);
                if (linkId.IsNull)
                {
                    statusMessage = "Р СњР Вµ РЎС“Р Т‘Р В°Р В»Р С•РЎРѓРЎРЉ РЎРѓР С•Р В·Р Т‘Р В°РЎвЂљРЎРЉ DataLink Р Т‘Р В»РЎРЏ Excel.";
                    return false;
                }

                BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                var table = new Table
                {
                    TableStyle = doc.Database.Tablestyle,
                    Position = position
                };
                table.SetSize(1, 1);
                ms.AppendEntity(table);
                tr.AddNewlyCreatedDBObject(table, true);

                table.Cells[0, 0].SetDataLink(linkId, true);
                table.Cells[0, 0].UpdateDataLink(UpdateDirection.SourceToData, (UpdateOption)ExcelLinkUpdateOptionsMask);
                table.GenerateLayout();

                tr.Commit();
            }

            statusMessage = $"Р РЋР Р†РЎРЏР В·Р В°Р Р…Р Р…Р В°РЎРЏ РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ Р В° Excel Р Р†РЎРѓРЎвЂљР В°Р Р†Р В»Р ВµР Р…Р В°: Р Т‘Р С‘Р В°Р С—Р В°Р В·Р С•Р Р… '{rangeToken}', DataLink '{dataLinkName}'.";
            return true;
        }
        catch (Exception ex)
        {
            statusMessage = $"Р С›РЎв‚¬Р С‘Р В±Р С”Р В° Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р С‘ РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…Р С•Р в„– РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎвЂ№ Excel: {ex.Message}";
            return false;
        }
        // END_BLOCK_TRY_INSERT_EXCEL_LINKED_TABLE
    }

    // START_CONTRACT: GetCachedOrLoadOutput
    //   PURPOSE: Retrieve cached or load output.
    //   INPUTS: { templatePath: string - method parameter }
    //   OUTPUTS: { IReadOnlyList<ExcelOutputRow> - result of retrieve cached or load output }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: GetCachedOrLoadOutput

    public IReadOnlyList<ExcelOutputRow> GetCachedOrLoadOutput(string templatePath)
    {
        // START_BLOCK_GET_CACHED_OR_LOAD_OUTPUT
        string key = NormalizePath(templatePath);
        if (OutputCache.TryGetValue(key, out CachedExcelOutput? cached))
        {
            _log.Write($"Р ВРЎРѓР С—Р С•Р В»РЎРЉР В·Р С•Р Р†Р В°Р Р… Р С”РЎРЊРЎв‚¬ Excel OUTPUT. Р РЋРЎвЂљРЎР‚Р С•Р С”: {cached.Rows.Count}.");
            return cached.Rows;
        }

        return FromExcelOutput(templatePath);
        // END_BLOCK_GET_CACHED_OR_LOAD_OUTPUT
    }

    // START_CONTRACT: TryGetCachedOutput
    //   PURPOSE: Attempt to execute get cached output.
    //   INPUTS: { templatePath: string - method parameter; rows: out IReadOnlyList<ExcelOutputRow> - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute get cached output }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: TryGetCachedOutput

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

    // START_CONTRACT: ClearCachedOutput
    //   PURPOSE: Clear cached output.
    //   INPUTS: { templatePath: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: ClearCachedOutput

    public void ClearCachedOutput(string templatePath)
    {
        // START_BLOCK_CLEAR_CACHED_OUTPUT
        _ = OutputCache.Remove(NormalizePath(templatePath));
        // END_BLOCK_CLEAR_CACHED_OUTPUT
    }

    // START_CONTRACT: ResolvePreferredExcelRangeToken
    //   PURPOSE: Resolve preferred excel range token.
    //   INPUTS: { workbookPath: string - method parameter }
    //   OUTPUTS: { string - textual result for resolve preferred excel range token }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: ResolvePreferredExcelRangeToken

    private static string ResolvePreferredExcelRangeToken(string workbookPath)
    {
        // START_BLOCK_RESOLVE_PREFERRED_EXCEL_RANGE_TOKEN
        try
        {
            using var workbook = new XLWorkbook(workbookPath);
            var acadRangeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (NamedRangeBinding definedName in EnumerateWorkbookRangeBindings(workbook))
            {
                if (string.Equals(definedName.WorksheetName, ExcelAcadSheetName, StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(definedName.Name))
                {
                    acadRangeNames.Add(definedName.Name);
                }
            }

            if (acadRangeNames.Contains(ExcelPrimaryRangeName))
            {
                return acadRangeNames.First(x => string.Equals(x, ExcelPrimaryRangeName, StringComparison.OrdinalIgnoreCase));
            }

            string? firstShieldRange = acadRangeNames
                .Where(name => name.StartsWith("Р В©Р С‘РЎвЂљ_", StringComparison.OrdinalIgnoreCase))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(firstShieldRange))
            {
                return firstShieldRange;
            }

            if (acadRangeNames.Contains(ExcelFallbackRangeName))
            {
                return acadRangeNames.First(x => string.Equals(x, ExcelFallbackRangeName, StringComparison.OrdinalIgnoreCase));
            }
        }
        catch
        {
            // If workbook cannot be parsed, use explicit fallback range.
        }

        return ExcelFallbackRangeAddress;
        // END_BLOCK_RESOLVE_PREFERRED_EXCEL_RANGE_TOKEN
    }

    // START_CONTRACT: EnumerateWorkbookRangeBindings
    //   PURPOSE: Enumerate workbook range bindings.
    //   INPUTS: { workbook: XLWorkbook - method parameter }
    //   OUTPUTS: { IEnumerable<NamedRangeBinding> - result of enumerate workbook range bindings }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: EnumerateWorkbookRangeBindings

    private static IEnumerable<NamedRangeBinding> EnumerateWorkbookRangeBindings(XLWorkbook workbook)
    {
        // START_BLOCK_ENUMERATE_WORKBOOK_RANGE_BINDINGS
        object? namesContainer = TryGetPropertyValue(workbook, "DefinedNames")
            ?? TryGetPropertyValue(workbook, "NamedRanges");
        if (namesContainer is null)
        {
            yield break;
        }

        foreach (object definedName in EnumerateObjects(namesContainer))
        {
            string name = ReadStringProperty(definedName, "Name");
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            object? rangesContainer = TryGetPropertyValue(definedName, "Ranges")
                ?? TryGetPropertyValue(definedName, "NamedRanges")
                ?? TryGetPropertyValue(definedName, "Range");
            if (rangesContainer is null)
            {
                continue;
            }

            foreach (object range in EnumerateObjects(rangesContainer))
            {
                string worksheetName = ReadStringProperty(TryGetPropertyValue(range, "Worksheet"), "Name");
                if (!string.IsNullOrWhiteSpace(worksheetName))
                {
                    yield return new NamedRangeBinding(name, worksheetName);
                }
            }
        }
        // END_BLOCK_ENUMERATE_WORKBOOK_RANGE_BINDINGS
    }

    // START_CONTRACT: EnumerateObjects
    //   PURPOSE: Enumerate objects.
    //   INPUTS: { candidate: object - method parameter }
    //   OUTPUTS: { IEnumerable<object> - result of enumerate objects }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: EnumerateObjects

    private static IEnumerable<object> EnumerateObjects(object candidate)
    {
        // START_BLOCK_ENUMERATE_OBJECTS_FOR_RANGE_BINDINGS
        if (candidate is string)
        {
            yield break;
        }

        if (candidate is IEnumerable enumerable)
        {
            foreach (object? item in enumerable)
            {
                if (item is not null)
                {
                    yield return item;
                }
            }

            yield break;
        }

        yield return candidate;
        // END_BLOCK_ENUMERATE_OBJECTS_FOR_RANGE_BINDINGS
    }

    // START_CONTRACT: TryGetPropertyValue
    //   PURPOSE: Attempt to execute get property value.
    //   INPUTS: { instance: object? - method parameter; propertyName: string - method parameter }
    //   OUTPUTS: { object? - result of attempt to execute get property value }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: TryGetPropertyValue

    private static object? TryGetPropertyValue(object? instance, string propertyName)
    {
        // START_BLOCK_TRY_GET_PROPERTY_VALUE_FOR_RANGE_BINDINGS
        if (instance is null)
        {
            return null;
        }

        try
        {
            PropertyInfo? property = instance.GetType().GetProperty(
                propertyName,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (property is null || !property.CanRead)
            {
                return null;
            }

            return property.GetValue(instance);
        }
        catch
        {
            return null;
        }
        // END_BLOCK_TRY_GET_PROPERTY_VALUE_FOR_RANGE_BINDINGS
    }

    // START_CONTRACT: ReadStringProperty
    //   PURPOSE: Read string property.
    //   INPUTS: { instance: object? - method parameter; propertyName: string - method parameter }
    //   OUTPUTS: { string - textual result for read string property }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: ReadStringProperty

    private static string ReadStringProperty(object? instance, string propertyName)
    {
        // START_BLOCK_READ_STRING_PROPERTY_FOR_RANGE_BINDINGS
        object? value = TryGetPropertyValue(instance, propertyName);
        return value?.ToString()?.Trim() ?? string.Empty;
        // END_BLOCK_READ_STRING_PROPERTY_FOR_RANGE_BINDINGS
    }

    // START_CONTRACT: BuildExcelDataLinkName
    //   PURPOSE: Build excel data link name.
    //   INPUTS: { workbookPath: string - method parameter; rangeToken: string - method parameter }
    //   OUTPUTS: { string - textual result for build excel data link name }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: BuildExcelDataLinkName

    private static string BuildExcelDataLinkName(string workbookPath, string rangeToken)
    {
        // START_BLOCK_BUILD_EXCEL_DATALINK_NAME
        string filePart = Path.GetFileNameWithoutExtension(workbookPath);
        string rawName = $"ELTOOLS_{filePart}_{rangeToken}";
        var normalized = new string(rawName
            .Select(ch => char.IsLetterOrDigit(ch) || ch is '_' or '-'
                ? ch
                : '_')
            .ToArray());

        return normalized.Length <= 120
            ? normalized
            : normalized[..120];
        // END_BLOCK_BUILD_EXCEL_DATALINK_NAME
    }

    // START_CONTRACT: BuildWorkbookConnectionPath
    //   PURPOSE: Build workbook connection path.
    //   INPUTS: { drawingPath: string - method parameter; workbookPath: string - method parameter }
    //   OUTPUTS: { string - textual result for build workbook connection path }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: BuildWorkbookConnectionPath

    private static string BuildWorkbookConnectionPath(string drawingPath, string workbookPath)
    {
        // START_BLOCK_BUILD_WORKBOOK_CONNECTION_PATH
        string normalizedWorkbookPath = NormalizePath(workbookPath);
        if (string.IsNullOrWhiteSpace(drawingPath) || !Path.IsPathRooted(drawingPath))
        {
            return normalizedWorkbookPath;
        }

        string drawingDirectory = Path.GetDirectoryName(NormalizePath(drawingPath)) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(drawingDirectory))
        {
            return normalizedWorkbookPath;
        }

        string relativePath = Path.GetRelativePath(drawingDirectory, normalizedWorkbookPath).Replace('/', '\\');
        if (!relativePath.StartsWith(".\\", StringComparison.OrdinalIgnoreCase)
            && !relativePath.StartsWith("..\\", StringComparison.OrdinalIgnoreCase))
        {
            relativePath = $".\\{relativePath}";
        }

        return relativePath;
        // END_BLOCK_BUILD_WORKBOOK_CONNECTION_PATH
    }

    // START_CONTRACT: EnsureExcelDataLink
    //   PURPOSE: Ensure excel data link.
    //   INPUTS: { tr: Transaction - method parameter; db: Database - method parameter; dataLinkName: string - method parameter; connectionString: string - method parameter }
    //   OUTPUTS: { ObjectId - result of ensure excel data link }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: EnsureExcelDataLink

    private static ObjectId EnsureExcelDataLink(
        Transaction tr,
        Database db,
        string dataLinkName,
        string connectionString)
    {
        // START_BLOCK_ENSURE_EXCEL_DATALINK
        DataLinkManager manager = db.DataLinkManager;
        ObjectId existingLinkId = ObjectId.Null;
        try
        {
            existingLinkId = manager.GetDataLink(dataLinkName);
        }
        catch
        {
            existingLinkId = ObjectId.Null;
        }

        if (!existingLinkId.IsNull)
        {
            if (tr.GetObject(existingLinkId, OpenMode.ForWrite, false) is DataLink existingLink)
            {
                existingLink.DataAdapterId = ExcelDataAdapterId;
                existingLink.ConnectionString = connectionString;
                existingLink.UpdateOption = ExcelLinkUpdateOptionsMask;
                existingLink.DataLinkOption = DataLinkOption.PersistCache;
                return existingLinkId;
            }
        }

        var link = new DataLink
        {
            Name = dataLinkName,
            DataAdapterId = ExcelDataAdapterId,
            ConnectionString = connectionString,
            UpdateOption = ExcelLinkUpdateOptionsMask,
            DataLinkOption = DataLinkOption.PersistCache
        };

        return manager.AddDataLink(link);
        // END_BLOCK_ENSURE_EXCEL_DATALINK
    }

    // START_CONTRACT: NormalizePath
    //   PURPOSE: Normalize path.
    //   INPUTS: { path: string - method parameter }
    //   OUTPUTS: { string - textual result for normalize path }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: NormalizePath

    private static string NormalizePath(string path)
    {
        // START_BLOCK_NORMALIZE_PATH
        return Path.GetFullPath(path);
        // END_BLOCK_NORMALIZE_PATH
    }

    // START_CONTRACT: LogWorkbookReadDiagnostics
    //   PURPOSE: Log workbook read diagnostics.
    //   INPUTS: { diagnostics: CsvExcelGateway.WorkbookReadDiagnostics - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXPORT
    // END_CONTRACT: LogWorkbookReadDiagnostics

    private void LogWorkbookReadDiagnostics(CsvExcelGateway.WorkbookReadDiagnostics diagnostics)
    {
        // START_BLOCK_LOG_WORKBOOK_READ_DIAGNOSTICS
        string sheets = diagnostics.SheetNames.Count == 0
            ? "<Р В»Р С‘РЎРѓРЎвЂљРЎвЂ№ Р Р…Р Вµ Р С•Р В±Р Р…Р В°РЎР‚РЎС“Р В¶Р ВµР Р…РЎвЂ№>"
            : string.Join(", ", diagnostics.SheetNames);
        _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: Р В»Р С‘РЎРѓРЎвЂљРЎвЂ№ Р С”Р Р…Р С‘Р С–Р С‘ = {sheets}");

        string selectedSheet = string.IsNullOrWhiteSpace(diagnostics.SelectedSheetName)
            ? "<Р В»Р С‘РЎРѓРЎвЂљ 'Р вЂ™ Р С’Р С”Р В°Р Т‘' Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р…>"
            : diagnostics.SelectedSheetName;
        _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: Р Р†РЎвЂ№Р В±РЎР‚Р В°Р Р…Р Р…РЎвЂ№Р в„– Р В»Р С‘РЎРѓРЎвЂљ = {selectedSheet}");

        if (!string.IsNullOrWhiteSpace(diagnostics.ImportWindow))
        {
            _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: Р С•Р С”Р Р…Р С• Р С‘Р СР С—Р С•РЎР‚РЎвЂљР В° = {diagnostics.ImportWindow}");
        }

        if (!string.IsNullOrWhiteSpace(diagnostics.ColumnMap))
        {
            _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: {diagnostics.ColumnMap}; data_start={diagnostics.DataStartRowNumber}; data_end={diagnostics.DataEndRowNumber}");
        }

        if (diagnostics.PreviewRows.Count > 0)
        {
            _log.Write("EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: Р С—Р ВµРЎР‚Р Р†РЎвЂ№Р Вµ 5 РЎРѓРЎвЂљРЎР‚Р С•Р С” Р В»Р С‘РЎРѓРЎвЂљР В°:");
            foreach (string previewLine in diagnostics.PreviewRows.Take(5))
            {
                _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: {previewLine}");
            }
        }

        if (diagnostics.ValidationIssues.Count > 0)
        {
            _log.Write("EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: Р С—РЎР‚Р С‘РЎвЂЎР С‘Р Р…РЎвЂ№ Р С•РЎвЂљР В±РЎР‚Р В°Р С”Р С•Р Р†Р С”Р С‘/Р С•РЎв‚¬Р С‘Р В±Р С•Р С”:");
            const int issueLimit = 40;
            foreach (string issue in diagnostics.ValidationIssues.Take(issueLimit))
            {
                _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: {issue}");
            }

            int hiddenIssues = Math.Max(0, diagnostics.ValidationIssues.Count - issueLimit) + diagnostics.SuppressedValidationIssueCount;
            if (hiddenIssues > 0)
            {
                _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: Р ВµРЎвЂ°Р Вµ {hiddenIssues} Р С—РЎР‚Р С‘РЎвЂЎР С‘Р Р… РЎРѓР С”РЎР‚РЎвЂ№РЎвЂљР С•.");
            }
        }

        if (!string.IsNullOrWhiteSpace(diagnostics.FatalError))
        {
            _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL DEBUG: Р С”РЎР‚Р С‘РЎвЂљР С‘РЎвЂЎР ВµРЎРѓР С”Р В°РЎРЏ Р С•РЎв‚¬Р С‘Р В±Р С”Р В° РЎвЂЎРЎвЂљР ВµР Р…Р С‘РЎРЏ Р С”Р Р…Р С‘Р С–Р С‘: {diagnostics.FatalError}");
        }
        // END_BLOCK_LOG_WORKBOOK_READ_DIAGNOSTICS
    }

    private readonly record struct NamedRangeBinding(string Name, string WorksheetName);

    private sealed record CachedExcelOutput(DateTime ImportedAtUtc, IReadOnlyList<ExcelOutputRow> Rows);
}