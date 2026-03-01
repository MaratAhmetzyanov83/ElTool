// FILE: src/Integrations/ExcelGateway.cs
// VERSION: 1.2.3
// START_MODULE_CONTRACT
//   PURPOSE: Define abstract gateway for Excel template I/O and provide CSV plus direct workbook fallback import.
//   SCOPE: INPUT sheet export and OUTPUT import from CSV or xlsx worksheet.
//   DEPENDS: M-CONFIG, M-MODELS
//   LINKS: M-EXCEL-GATEWAY, M-CONFIG, M-MODELS
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   IExcelGateway - Abstraction for excel read/write operations.
//   CsvExcelGateway - Writes INPUT csv and reads OUTPUT from csv or from workbook sheet.
// END_MODULE_MAP
//
// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.2.2 - Added range-aware workbook import: prefer named range 'Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎС™Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ_Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В·Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В ' on sheet 'Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РЎС› Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’В' and apply canonical column mapping (D/E/K/M with G note) before generic header scan fallback.
// END_CHANGE_SUMMARY

using System.Collections;
using System.Globalization;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;
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
    private const string AcadSheetName = "Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РЎС› Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’В";
    private const string AcadImportRangeName = "Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎС™Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ_Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В·Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В ";
    private const int CanonicalShieldColumn = 4; // D
    private const int CanonicalLineColumn = 5; // E
    private const int CanonicalNoteColumn = 7; // G
    private const int CanonicalBreakerColumn = 11; // K
    private const int CanonicalCableColumn = 13; // M
    private const int HeaderScanRowLimit = 200;
    private const int HeaderScanColumnLimit = 80;
    private const int PreviewRowCount = 5;
    private const int PreviewColumnCount = 16;
    private static readonly string[] WorkbookExtensions = [".xlsx", ".xlsm", ".xltx", ".xltm"];
    private static readonly string[] ShieldHeaderTokens = ["Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ"];
    private static readonly string[] LineHeaderTokens = ["Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ", "Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В"];
    private static readonly string[] CableHeaderTokens = ["Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°"];
    private static readonly string[] BreakerHeaderTokens = ["Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ", "Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ"];
    private static readonly string[] NoteHeaderTokens = ["Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№", "Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ", "Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В·Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ"];

    private WorkbookReadDiagnostics? _lastWorkbookReadDiagnostics;

    // START_CONTRACT: GetLastWorkbookReadDiagnostics
    //   PURPOSE: Retrieve last workbook read diagnostics.
    //   INPUTS: none
    //   OUTPUTS: { WorkbookReadDiagnostics? - result of retrieve last workbook read diagnostics }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: GetLastWorkbookReadDiagnostics

    public WorkbookReadDiagnostics? GetLastWorkbookReadDiagnostics()
    {
        // START_BLOCK_GET_LAST_WORKBOOK_READ_DIAGNOSTICS
        return _lastWorkbookReadDiagnostics;
        // END_BLOCK_GET_LAST_WORKBOOK_READ_DIAGNOSTICS
    }

    // START_CONTRACT: WriteInputRows
    //   PURPOSE: Write input rows.
    //   INPUTS: { templatePath: string - method parameter; rows: IReadOnlyList<ExcelInputRow> - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: WriteInputRows

    public void WriteInputRows(string templatePath, IReadOnlyList<ExcelInputRow> rows)
    {
        // START_BLOCK_WRITE_INPUT_ROWS
        string directory = Path.GetDirectoryName(templatePath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        string output = Path.Combine(string.IsNullOrWhiteSpace(directory) ? "." : directory, $"{fileName}.INPUT.csv");

        var lines = new List<string>
        {
            string.Join(Separator, new[]
            {
                "Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРІР‚С”","Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†РІР‚С™Р’В¬Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћ","Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎвЂќР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎвЂќР В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРІР‚С”Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¬_Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РЎС›Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ","Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР Р‹Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎвЂєР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎвЂє","Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎС™Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћ_Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’В","Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎС™Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћ_Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎвЂќР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРІР‚С”Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎвЂќР В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎвЂќР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›_Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’В","Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎС™Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћ_Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎвЂќР В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚Сњ_Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’В","Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎС™Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћ_Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРІР‚С”Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎвЂќР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР Р‹Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›_Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’В","Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¤Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћ","Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРІР‚С”Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџ_Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†РІР‚С™Р’В¬Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В«"
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

    // START_CONTRACT: ReadOutputRows
    //   PURPOSE: Read output rows.
    //   INPUTS: { templatePath: string - method parameter }
    //   OUTPUTS: { IReadOnlyList<ExcelOutputRow> - result of read output rows }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ReadOutputRows

    public IReadOnlyList<ExcelOutputRow> ReadOutputRows(string templatePath)
    {
        // START_BLOCK_READ_OUTPUT_ROWS
        _lastWorkbookReadDiagnostics = null;

        string directory = Path.GetDirectoryName(templatePath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        string csvPath = Path.Combine(string.IsNullOrWhiteSpace(directory) ? "." : directory, $"{fileName}.OUTPUT.csv");
        List<ExcelOutputRow>? csvRows = null;

        if (File.Exists(csvPath))
        {
            csvRows = ReadOutputRowsFromCsv(csvPath).ToList();
            if (csvRows.Count > 0)
            {
                return csvRows;
            }
        }

        if (!File.Exists(templatePath))
        {
            return csvRows ?? [];
        }

        string extension = Path.GetExtension(templatePath);
        if (!WorkbookExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            return csvRows ?? [];
        }

        IReadOnlyList<ExcelOutputRow> workbookRows = ReadOutputRowsFromWorkbook(templatePath);
        if (workbookRows.Count > 0 || csvRows is null)
        {
            return workbookRows;
        }

        return csvRows;
        // END_BLOCK_READ_OUTPUT_ROWS
    }

    // START_CONTRACT: ReadOutputRowsFromCsv
    //   PURPOSE: Read output rows from csv.
    //   INPUTS: { inputPath: string - method parameter }
    //   OUTPUTS: { IReadOnlyList<ExcelOutputRow> - result of read output rows from csv }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ReadOutputRowsFromCsv

    private static IReadOnlyList<ExcelOutputRow> ReadOutputRowsFromCsv(string inputPath)
    {
        // START_BLOCK_READ_OUTPUT_ROWS_FROM_CSV
        var rows = new List<ExcelOutputRow>();
        string[] lines = File.ReadAllLines(inputPath, Encoding.UTF8);
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
        // END_BLOCK_READ_OUTPUT_ROWS_FROM_CSV
    }

    // START_CONTRACT: ReadOutputRowsFromWorkbook
    //   PURPOSE: Read output rows from workbook.
    //   INPUTS: { workbookPath: string - method parameter }
    //   OUTPUTS: { IReadOnlyList<ExcelOutputRow> - result of read output rows from workbook }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ReadOutputRowsFromWorkbook

    private IReadOnlyList<ExcelOutputRow> ReadOutputRowsFromWorkbook(string workbookPath)
    {
        // START_BLOCK_READ_OUTPUT_ROWS_FROM_WORKBOOK
        var diagnostics = new WorkbookReadDiagnostics(workbookPath);
        _lastWorkbookReadDiagnostics = diagnostics;

        try
        {
            using var stream = new FileStream(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);

            diagnostics.SheetNames.AddRange(workbook.Worksheets.Select(sheet => sheet.Name));

            IXLWorksheet? sheet = FindAcadSheet(workbook);
            if (sheet is null)
            {
                diagnostics.AddValidationIssue("Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ 'Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РЎС› Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’В' Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р вЂ Р Р†Р вЂљРЎвЂєР Р†Р вЂљРІР‚СљР В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦.");
                return [];
            }

            diagnostics.SelectedSheetName = sheet.Name;
            int lastRow = GetLastRowNumber(sheet);
            int lastColumn = GetLastColumnNumber(sheet);
            diagnostics.PreviewRows.AddRange(BuildPreviewRows(sheet, PreviewRowCount, Math.Max(lastColumn, 8)));

            if (lastRow <= 0 || lastColumn <= 0)
            {
                diagnostics.AddValidationIssue("Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РЎС›Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р вЂ™Р’В Р В Р вЂ Р Р†Р вЂљРЎвЂєР Р†Р вЂљРІР‚Сљ Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р вЂ™Р’В¦.");
                return [];
            }

            ImportWindow importWindow = new(1, lastRow, 1, lastColumn, string.Empty);
            if (TryResolveImportWindow(workbook, sheet, lastRow, lastColumn, out ImportWindow rangeWindow))
            {
                importWindow = rangeWindow;
                diagnostics.ImportWindow = importWindow.DebugLabel;
            }
            else
            {
                diagnostics.ImportWindow = $"used_range=A1:{ToColumnName(lastColumn)}{lastRow}";
            }

            bool useCanonicalMap = TryBuildCanonicalAcadHeaderMap(importWindow, out HeaderMap headerMap);
            if (!useCanonicalMap
                && !TryResolveHeaderMap(sheet, importWindow.MinRow, importWindow.MaxRow, importWindow.MinColumn, importWindow.MaxColumn, out headerMap)
                && !TryResolveHeaderMap(sheet, lastRow, lastColumn, out headerMap))
            {
                diagnostics.AddValidationIssue("Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р вЂ Р Р†Р вЂљРЎвЂєР Р†Р вЂљРІР‚СљР В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В·Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В  Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС› Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРІвЂћвЂ“Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’В 'Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ', 'Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ', 'Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°'.");
                return [];
            }

            diagnostics.HeaderRowNumber = headerMap.RowNumber;
            diagnostics.DataStartRowNumber = Math.Max(headerMap.RowNumber + 1, importWindow.MinRow);
            diagnostics.DataEndRowNumber = importWindow.MaxRow;
            diagnostics.ColumnMap = useCanonicalMap
                ? $"{headerMap.ToDebugString()}; source=canonical_acad_columns"
                : headerMap.ToDebugString();

            if (diagnostics.DataStartRowNumber > diagnostics.DataEndRowNumber)
            {
                diagnostics.AddValidationIssue("Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В·Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В  Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљ Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р вЂ™Р’В¦.");
                return [];
            }

            var shadowRows = new List<WorkbookShadowRow>();
            for (int row = diagnostics.DataStartRowNumber; row <= diagnostics.DataEndRowNumber; row++)
            {
                try
                {
                    if (TryBuildShadowRow(sheet, row, headerMap, out WorkbookShadowRow shadowRow, out string? rejectionReason))
                    {
                        shadowRows.Add(shadowRow);
                    }
                    else if (!string.IsNullOrWhiteSpace(rejectionReason))
                    {
                        diagnostics.AddValidationIssue(rejectionReason);
                    }
                }
                catch (Exception ex)
                {
                    diagnostics.AddValidationIssue($"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р вЂ Р Р†Р вЂљРЎв„ўР вЂ™Р’В¬Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В ({ex.Message}).");
                }
            }

            List<ExcelOutputRow> rows = shadowRows.Select(ToExcelOutputRow).ToList();
            diagnostics.AcceptedRows = rows.Count;
            if (rows.Count == 0 && diagnostics.ValidationIssues.Count == 0)
            {
                diagnostics.AddValidationIssue("Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В» Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р вЂ™Р’В¦ Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІР‚С”Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р вЂ™Р’В Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р вЂ™Р’В¦ Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р вЂ™Р’В¦.");
            }

            return rows;
        }
        catch (Exception ex)
        {
            diagnostics.FatalError = ex.Message;
            diagnostics.AddValidationIssue($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎвЂќР В Р’В Р В Р вЂ№Р В Р вЂ Р Р†Р вЂљРЎв„ўР вЂ™Р’В¬Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ/Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В: {ex.Message}");
            return [];
        }
        
        // END_BLOCK_READ_OUTPUT_ROWS_FROM_WORKBOOK
    }
    // START_CONTRACT: TryResolveImportWindow
    //   PURPOSE: Attempt to execute resolve import window.
    //   INPUTS: { workbook: XLWorkbook - method parameter; sheet: IXLWorksheet - method parameter; lastRow: int - method parameter; lastColumn: int - method parameter; window: out ImportWindow - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute resolve import window }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: TryResolveImportWindow

    private static bool TryResolveImportWindow(
        XLWorkbook workbook,
        IXLWorksheet sheet,
        int lastRow,
        int lastColumn,
        out ImportWindow window)
    {
        // START_BLOCK_TRY_RESOLVE_IMPORT_WINDOW
        string expectedName = NormalizeForLookup(AcadImportRangeName);
        string expectedSheet = NormalizeForLookup(sheet.Name);
        foreach (NamedRangeWindow namedRange in EnumerateWorkbookNamedRanges(workbook))
        {
            if (!string.Equals(NormalizeForLookup(namedRange.Name), expectedName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!string.Equals(NormalizeForLookup(namedRange.WorksheetName), expectedSheet, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            int minRow = Math.Max(1, namedRange.MinRow);
            int maxRow = Math.Min(lastRow, namedRange.MaxRow);
            int minColumn = Math.Max(1, namedRange.MinColumn);
            int maxColumn = Math.Min(lastColumn, namedRange.MaxColumn);

            if (minRow > maxRow || minColumn > maxColumn)
            {
                continue;
            }

            string label = string.IsNullOrWhiteSpace(namedRange.Address)
                ? namedRange.Name
                : $"{namedRange.Name}={namedRange.Address}";
            window = new ImportWindow(minRow, maxRow, minColumn, maxColumn, label);
            return true;
        }

        window = new ImportWindow(1, lastRow, 1, lastColumn, string.Empty);
        return false;
        // END_BLOCK_TRY_RESOLVE_IMPORT_WINDOW
    }

    // START_CONTRACT: TryBuildCanonicalAcadHeaderMap
    //   PURPOSE: Attempt to execute build canonical acad header map.
    //   INPUTS: { window: ImportWindow - method parameter; headerMap: out HeaderMap - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute build canonical acad header map }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: TryBuildCanonicalAcadHeaderMap

    private static bool TryBuildCanonicalAcadHeaderMap(ImportWindow window, out HeaderMap headerMap)
    {
        // START_BLOCK_TRY_BUILD_CANONICAL_ACAD_HEADER_MAP
        int headerRow = Math.Max(1, window.MinRow - 1);
        if (!window.ContainsColumn(CanonicalShieldColumn)
            || !window.ContainsColumn(CanonicalLineColumn)
            || !window.ContainsColumn(CanonicalCableColumn))
        {
            headerMap = HeaderMap.Empty;
            return false;
        }

        headerMap = new HeaderMap(
            headerRow,
            CanonicalShieldColumn,
            CanonicalLineColumn,
            CanonicalCableColumn,
            window.ContainsColumn(CanonicalBreakerColumn) ? CanonicalBreakerColumn : 0,
            window.ContainsColumn(CanonicalNoteColumn) ? CanonicalNoteColumn : 0);
        return true;
        // END_BLOCK_TRY_BUILD_CANONICAL_ACAD_HEADER_MAP
    }

    // START_CONTRACT: TryResolveHeaderMap
    //   PURPOSE: Attempt to execute resolve header map.
    //   INPUTS: { sheet: IXLWorksheet - method parameter; lastRow: int - method parameter; lastColumn: int - method parameter; headerMap: out HeaderMap - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute resolve header map }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: TryResolveHeaderMap

    private static bool TryResolveHeaderMap(IXLWorksheet sheet, int lastRow, int lastColumn, out HeaderMap headerMap)
    {
        // START_BLOCK_TRY_RESOLVE_HEADER_MAP
        return TryResolveHeaderMap(
            sheet,
            1,
            Math.Min(lastRow, HeaderScanRowLimit),
            1,
            Math.Min(lastColumn, HeaderScanColumnLimit),
            out headerMap);
        // END_BLOCK_TRY_RESOLVE_HEADER_MAP
    }

    // START_CONTRACT: TryResolveHeaderMap
    //   PURPOSE: Attempt to execute resolve header map.
    //   INPUTS: { sheet: IXLWorksheet - method parameter; minRow: int - method parameter; maxRow: int - method parameter; minColumn: int - method parameter; maxColumn: int - method parameter; headerMap: out HeaderMap - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute resolve header map }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: TryResolveHeaderMap

    private static bool TryResolveHeaderMap(
        IXLWorksheet sheet,
        int minRow,
        int maxRow,
        int minColumn,
        int maxColumn,
        out HeaderMap headerMap)
    {
        // START_BLOCK_TRY_RESOLVE_HEADER_MAP_WINDOW
        int rowStart = Math.Max(1, minRow);
        int rowLimit = Math.Max(rowStart, Math.Min(maxRow, HeaderScanRowLimit));
        int columnStart = Math.Max(1, minColumn);
        int columnLimit = Math.Max(columnStart, Math.Min(maxColumn, HeaderScanColumnLimit));

        for (int row = rowStart; row <= rowLimit; row++)
        {
            int shieldColumn = 0;
            int lineColumn = 0;
            int cableColumn = 0;
            int breakerColumn = 0;
            int noteColumn = 0;

            for (int column = columnStart; column <= columnLimit; column++)
            {
                string header = ReadHeaderCellText(sheet, row, column);
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                if (shieldColumn == 0 && ContainsAnyToken(header, ShieldHeaderTokens))
                {
                    shieldColumn = column;
                }

                if (lineColumn == 0 && ContainsAnyToken(header, LineHeaderTokens))
                {
                    lineColumn = column;
                }

                if (cableColumn == 0 && ContainsAnyToken(header, CableHeaderTokens))
                {
                    cableColumn = column;
                }

                if (breakerColumn == 0 && ContainsAnyToken(header, BreakerHeaderTokens))
                {
                    breakerColumn = column;
                }

                if (noteColumn == 0 && ContainsAnyToken(header, NoteHeaderTokens))
                {
                    noteColumn = column;
                }
            }

            if (shieldColumn > 0 && lineColumn > 0 && cableColumn > 0)
            {
                headerMap = new HeaderMap(row, shieldColumn, lineColumn, cableColumn, breakerColumn, noteColumn);
                return true;
            }
        }

        headerMap = HeaderMap.Empty;
        return false;
        // END_BLOCK_TRY_RESOLVE_HEADER_MAP_WINDOW
    }

    // START_CONTRACT: TryBuildShadowRow
    //   PURPOSE: Attempt to execute build shadow row.
    //   INPUTS: { sheet: IXLWorksheet - method parameter; row: int - method parameter; headerMap: HeaderMap - method parameter; shadowRow: out WorkbookShadowRow - method parameter; rejectionReason: out string? - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute build shadow row }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: TryBuildShadowRow

    private static bool TryBuildShadowRow(
        IXLWorksheet sheet,
        int row,
        HeaderMap headerMap,
        out WorkbookShadowRow shadowRow,
        out string? rejectionReason)
    {
        // START_BLOCK_TRY_BUILD_SHADOW_ROW
        shadowRow = WorkbookShadowRow.Empty;
        rejectionReason = null;

        CellReadResult shieldResult = ReadCellValue(sheet, row, headerMap.ShieldColumn, "Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ");
        CellReadResult lineResult = ReadCellValue(sheet, row, headerMap.LineColumn, "Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ");
        CellReadResult cableResult = ReadCellValue(sheet, row, headerMap.CableColumn, "Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°");

        if (shieldResult.IsEmpty && lineResult.IsEmpty && cableResult.IsEmpty)
        {
            return false;
        }

        if (shieldResult.HasError)
        {
            rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: {shieldResult.ErrorMessage} ({GetCellAddress(row, headerMap.ShieldColumn)}).";
            return false;
        }

        if (lineResult.HasError)
        {
            rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: {lineResult.ErrorMessage} ({GetCellAddress(row, headerMap.LineColumn)}).";
            return false;
        }

        if (cableResult.HasError)
        {
            rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: {cableResult.ErrorMessage} ({GetCellAddress(row, headerMap.CableColumn)}).";
            return false;
        }

        string shield = NormalizeSingleLineText(shieldResult.Value);
        string line = NormalizeSingleLineText(lineResult.Value);
        string cable = NormalizeCableValue(cableResult.Value);
        if (!IsMeaningfulField(shield))
        {
            rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° 'Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ' Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ ({GetCellAddress(row, headerMap.ShieldColumn)}).";
            return false;
        }

        if (!IsMeaningfulField(line))
        {
            rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° 'Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ' Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ ({GetCellAddress(row, headerMap.LineColumn)}).";
            return false;
        }

        if (!IsMeaningfulField(cable))
        {
            rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° 'Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°' Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ ({GetCellAddress(row, headerMap.CableColumn)}).";
            return false;
        }

        if (string.Equals(line, "0", StringComparison.OrdinalIgnoreCase))
        {
            rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° 'Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ' Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ 0 ({GetCellAddress(row, headerMap.LineColumn)}).";
            return false;
        }

        string breaker = "QF";
        if (headerMap.BreakerColumn > 0)
        {
            CellReadResult breakerResult = ReadCellValue(sheet, row, headerMap.BreakerColumn, "Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ");
            if (breakerResult.HasError)
            {
                rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: {breakerResult.ErrorMessage} ({GetCellAddress(row, headerMap.BreakerColumn)}).";
                return false;
            }

            if (breakerResult.HasValue && IsMeaningfulField(breakerResult.Value))
            {
                breaker = NormalizeSingleLineText(breakerResult.Value);
            }
        }

        if (!IsMeaningfulField(breaker) || string.Equals(breaker, line, StringComparison.OrdinalIgnoreCase))
        {
            breaker = "QF";
        }

        string note = string.Empty;
        if (headerMap.NoteColumn > 0)
        {
            CellReadResult noteResult = ReadCellValue(sheet, row, headerMap.NoteColumn, "Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ");
            if (noteResult.HasError)
            {
                rejectionReason = $"Р В Р’В Р вЂ™Р’В Р В Р’В Р В РІР‚в„–Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° {row}: {noteResult.ErrorMessage} ({GetCellAddress(row, headerMap.NoteColumn)}).";
                return false;
            }

            if (noteResult.HasValue)
            {
                note = NormalizeSingleLineText(noteResult.Value);
            }
        }

        var cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ"] = shield,
            ["Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ"] = line,
            ["Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°"] = cable,
            ["Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ"] = breaker,
            ["Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ"] = note
        };

        shadowRow = new WorkbookShadowRow(row, cells);
        return true;
        // END_BLOCK_TRY_BUILD_SHADOW_ROW
    }

    // START_CONTRACT: ToExcelOutputRow
    //   PURPOSE: To excel output row.
    //   INPUTS: { shadowRow: WorkbookShadowRow - method parameter }
    //   OUTPUTS: { ExcelOutputRow - result of to excel output row }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ToExcelOutputRow

    private static ExcelOutputRow ToExcelOutputRow(WorkbookShadowRow shadowRow)
    {
        // START_BLOCK_TO_EXCEL_OUTPUT_ROW
        shadowRow.Cells.TryGetValue("Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ", out string? note);
        return new ExcelOutputRow(
            shadowRow.Cells["Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ"],
            shadowRow.Cells["Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ"],
            shadowRow.Cells["Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ"],
            string.Empty,
            shadowRow.Cells["Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°"],
            1,
            0,
            note);
        // END_BLOCK_TO_EXCEL_OUTPUT_ROW
    }

    // START_CONTRACT: ReadCellValue
    //   PURPOSE: Read cell value.
    //   INPUTS: { sheet: IXLWorksheet - method parameter; row: int - method parameter; column: int - method parameter; columnName: string - method parameter }
    //   OUTPUTS: { CellReadResult - result of read cell value }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ReadCellValue

    private static CellReadResult ReadCellValue(IXLWorksheet sheet, int row, int column, string columnName)
    {
        // START_BLOCK_READ_CELL_VALUE
        IXLCell cell;
        try
        {
            cell = sheet.Cell(row, column);
        }
        catch (Exception ex)
        {
            return CellReadResult.Error($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° '{columnName}' Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°: {ex.Message}");
        }

        bool isEmpty;
        try
        {
            isEmpty = cell.IsEmpty();
        }
        catch (Exception ex)
        {
            return CellReadResult.Error($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° '{columnName}' Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°: {ex.Message}");
        }

        if (isEmpty)
        {
            return CellReadResult.Empty();
        }

        bool hasFormula;
        try
        {
            hasFormula = cell.HasFormula;
        }
        catch
        {
            hasFormula = false;
        }

        if (hasFormula)
        {
            try
            {
                string cached = NormalizeCellText(cell.CachedValue.ToString());
                if (string.IsNullOrWhiteSpace(cached))
                {
                    return CellReadResult.Error($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° '{columnName}' Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІР‚С”Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™ Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В· Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС› Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В·Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ");
                }

                if (cached.StartsWith('#'))
                {
                    return CellReadResult.Error($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° '{columnName}' Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р вЂ Р Р†Р вЂљРЎв„ўР вЂ™Р’В¬Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™ Excel '{cached}'");
                }

                return CellReadResult.FromValue(cached);
            }
            catch (Exception ex)
            {
                return CellReadResult.Error($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° '{columnName}' Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІР‚С”Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™, Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС› Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р В Р вЂ№Р В Р’В Р В Р вЂ°Р В Р’В Р В Р вЂ№Р В Р вЂ Р Р†Р вЂљРЎв„ўР вЂ™Р’В¬ Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦ ({ex.Message})");
            }
        }

        XLDataType dataType;
        try
        {
            dataType = cell.DataType;
        }
        catch
        {
            dataType = XLDataType.Text;
        }

        if (dataType == XLDataType.Blank)
        {
            return CellReadResult.Empty();
        }

        if (dataType == XLDataType.Error)
        {
            string errorValue = NormalizeCellText(SafeReadCellString(cell));
            if (string.IsNullOrWhiteSpace(errorValue))
            {
                errorValue = "#ERROR";
            }

            return CellReadResult.Error($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° '{columnName}' Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р вЂ Р Р†Р вЂљРЎв„ўР вЂ™Р’В¬Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™ Excel '{errorValue}'");
        }

        if (!IsSupportedDataType(dataType))
        {
            return CellReadResult.Error($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° '{columnName}' Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р Р†РІР‚С›РІР‚вЂњР В Р’В Р вЂ™Р’В Р В Р вЂ Р Р†Р вЂљРЎвЂєР Р†Р вЂљРІР‚Сљ Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚Сњ '{dataType}'");
        }

        string value = NormalizeCellText(SafeReadCellString(cell));
        if (string.IsNullOrWhiteSpace(value))
        {
            return CellReadResult.Empty();
        }

        if (value.StartsWith('#'))
        {
            return CellReadResult.Error($"Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В° '{columnName}' Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В¶Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р В Р вЂ№Р В Р вЂ Р Р†Р вЂљРЎв„ўР вЂ™Р’В¬Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™ Excel '{value}'");
        }

        return CellReadResult.FromValue(value);
        // END_BLOCK_READ_CELL_VALUE
    }

    // START_CONTRACT: ReadHeaderCellText
    //   PURPOSE: Read header cell text.
    //   INPUTS: { sheet: IXLWorksheet - method parameter; row: int - method parameter; column: int - method parameter }
    //   OUTPUTS: { string - textual result for read header cell text }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ReadHeaderCellText

    private static string ReadHeaderCellText(IXLWorksheet sheet, int row, int column)
    {
        // START_BLOCK_READ_HEADER_CELL_TEXT
        try
        {
            IXLCell cell = sheet.Cell(row, column);
            if (cell.IsEmpty())
            {
                return string.Empty;
            }

            if (cell.HasFormula)
            {
                try
                {
                    string cached = NormalizeCellText(cell.CachedValue.ToString());
                    if (!string.IsNullOrWhiteSpace(cached))
                    {
                        return cached;
                    }
                }
                catch
                {
                    // Ignore and continue with raw string extraction.
                }
            }

            return NormalizeCellText(SafeReadCellString(cell));
        }
        catch
        {
            return string.Empty;
        }
        // END_BLOCK_READ_HEADER_CELL_TEXT
    }

    // START_CONTRACT: SafeReadCellString
    //   PURPOSE: Safe read cell string.
    //   INPUTS: { cell: IXLCell - method parameter }
    //   OUTPUTS: { string - textual result for safe read cell string }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: SafeReadCellString

    private static string SafeReadCellString(IXLCell cell)
    {
        // START_BLOCK_SAFE_READ_CELL_STRING
        string value = string.Empty;
        try
        {
            value = cell.GetFormattedString();
        }
        catch
        {
            value = string.Empty;
        }

        if (!string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        try
        {
            return cell.GetString();
        }
        catch
        {
            return string.Empty;
        }
        // END_BLOCK_SAFE_READ_CELL_STRING
    }

    // START_CONTRACT: IsSupportedDataType
    //   PURPOSE: Check whether supported data type.
    //   INPUTS: { dataType: XLDataType - method parameter }
    //   OUTPUTS: { bool - true when method can check whether supported data type }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: IsSupportedDataType

    private static bool IsSupportedDataType(XLDataType dataType)
    {
        // START_BLOCK_IS_SUPPORTED_DATA_TYPE
        return dataType is XLDataType.Text
            or XLDataType.Number
            or XLDataType.DateTime
            or XLDataType.TimeSpan
            or XLDataType.Boolean;
        // END_BLOCK_IS_SUPPORTED_DATA_TYPE
    }

    // START_CONTRACT: GetLastRowNumber
    //   PURPOSE: Retrieve last row number.
    //   INPUTS: { sheet: IXLWorksheet - method parameter }
    //   OUTPUTS: { int - result of retrieve last row number }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: GetLastRowNumber

    private static int GetLastRowNumber(IXLWorksheet sheet)
    {
        // START_BLOCK_GET_LAST_ROW_NUMBER
        try
        {
            return sheet.LastRowUsed()?.RowNumber() ?? 0;
        }
        catch
        {
            try
            {
                return sheet.LastRow().RowNumber();
            }
            catch
            {
                return 0;
            }
        }
        // END_BLOCK_GET_LAST_ROW_NUMBER
    }

    // START_CONTRACT: GetLastColumnNumber
    //   PURPOSE: Retrieve last column number.
    //   INPUTS: { sheet: IXLWorksheet - method parameter }
    //   OUTPUTS: { int - result of retrieve last column number }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: GetLastColumnNumber

    private static int GetLastColumnNumber(IXLWorksheet sheet)
    {
        // START_BLOCK_GET_LAST_COLUMN_NUMBER
        try
        {
            return sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        }
        catch
        {
            try
            {
                return sheet.LastColumn().ColumnNumber();
            }
            catch
            {
                return 0;
            }
        }
        // END_BLOCK_GET_LAST_COLUMN_NUMBER
    }

    // START_CONTRACT: FindAcadSheet
    //   PURPOSE: Find acad sheet.
    //   INPUTS: { workbook: XLWorkbook - method parameter }
    //   OUTPUTS: { IXLWorksheet? - result of find acad sheet }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: FindAcadSheet

    private static IXLWorksheet? FindAcadSheet(XLWorkbook workbook)
    {
        // START_BLOCK_FIND_ACAD_SHEET
        string expected = NormalizeForLookup(AcadSheetName);
        IXLWorksheet? exact = workbook.Worksheets.FirstOrDefault(sheet =>
            string.Equals(NormalizeForLookup(sheet.Name), expected, StringComparison.OrdinalIgnoreCase));
        if (exact is not null)
        {
            return exact;
        }

        return workbook.Worksheets.FirstOrDefault(sheet =>
            NormalizeForLookup(sheet.Name).Contains("Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РЎС›Р Р†Р вЂљР’В", StringComparison.OrdinalIgnoreCase));
        // END_BLOCK_FIND_ACAD_SHEET
    }

    // START_CONTRACT: BuildPreviewRows
    //   PURPOSE: Build preview rows.
    //   INPUTS: { sheet: IXLWorksheet - method parameter; rowCount: int - method parameter; columnCount: int - method parameter }
    //   OUTPUTS: { IReadOnlyList<string> - result of build preview rows }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: BuildPreviewRows

    private static IReadOnlyList<string> BuildPreviewRows(IXLWorksheet sheet, int rowCount, int columnCount)
    {
        // START_BLOCK_BUILD_PREVIEW_ROWS
        int rowsToRead = Math.Max(1, rowCount);
        int columnsToRead = Math.Max(1, Math.Min(columnCount, PreviewColumnCount));
        var preview = new List<string>(rowsToRead);
        for (int row = 1; row <= rowsToRead; row++)
        {
            var parts = new List<string>();
            for (int column = 1; column <= columnsToRead; column++)
            {
                string text = ReadHeaderCellText(sheet, row, column);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                parts.Add($"{ToColumnName(column)}{row}='{text}'");
            }

            preview.Add(parts.Count == 0
                ? $"R{row}: <Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІР‚СњР В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎС™Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СљР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›>"
                : $"R{row}: {string.Join("; ", parts)}");
        }

        return preview;
        // END_BLOCK_BUILD_PREVIEW_ROWS
    }

    // START_CONTRACT: GetCellAddress
    //   PURPOSE: Retrieve cell address.
    //   INPUTS: { row: int - method parameter; column: int - method parameter }
    //   OUTPUTS: { string - textual result for retrieve cell address }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: GetCellAddress

    private static string GetCellAddress(int row, int column)
    {
        // START_BLOCK_GET_CELL_ADDRESS
        return $"{ToColumnName(column)}{row}";
        // END_BLOCK_GET_CELL_ADDRESS
    }

    // START_CONTRACT: ToColumnName
    //   PURPOSE: To column name.
    //   INPUTS: { columnNumber: int - method parameter }
    //   OUTPUTS: { string - textual result for to column name }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ToColumnName

    private static string ToColumnName(int columnNumber)
    {
        // START_BLOCK_TO_COLUMN_NAME
        if (columnNumber <= 0)
        {
            return "?";
        }

        int current = columnNumber;
        var result = new StringBuilder();
        while (current > 0)
        {
            current--;
            result.Insert(0, (char)('A' + (current % 26)));
            current /= 26;
        }

        return result.ToString();
        // END_BLOCK_TO_COLUMN_NAME
    }

    // START_CONTRACT: ContainsAnyToken
    //   PURPOSE: Contains any token.
    //   INPUTS: { source: string - method parameter; tokens: IEnumerable<string> - method parameter }
    //   OUTPUTS: { bool - true when method can contains any token }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ContainsAnyToken

    private static bool ContainsAnyToken(string source, IEnumerable<string> tokens)
    {
        // START_BLOCK_CONTAINS_ANY_TOKEN
        string normalized = NormalizeForLookup(source);
        foreach (string token in tokens)
        {
            if (normalized.Contains(NormalizeForLookup(token), StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
        // END_BLOCK_CONTAINS_ANY_TOKEN
    }

    // START_CONTRACT: IsMeaningfulField
    //   PURPOSE: Check whether meaningful field.
    //   INPUTS: { value: string - method parameter }
    //   OUTPUTS: { bool - true when method can check whether meaningful field }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: IsMeaningfulField

    private static bool IsMeaningfulField(string value)
    {
        // START_BLOCK_IS_MEANINGFUL_FIELD
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (value.StartsWith('#') || value.StartsWith('='))
        {
            return false;
        }

        return !string.Equals(value, "Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(value, "Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ў Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’В", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(value, "Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(value, "Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°", StringComparison.OrdinalIgnoreCase);
        // END_BLOCK_IS_MEANINGFUL_FIELD
    }

    // START_CONTRACT: TryParseInt
    //   PURPOSE: Attempt to execute parse int.
    //   INPUTS: { raw: string? - method parameter }
    //   OUTPUTS: { int - result of attempt to execute parse int }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: TryParseInt

    private static int TryParseInt(string? raw)
    {
        // START_BLOCK_TRY_PARSE_INT
        return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : 0;
        // END_BLOCK_TRY_PARSE_INT
    }

    // START_CONTRACT: Escape
    //   PURPOSE: Escape.
    //   INPUTS: { value: string - method parameter }
    //   OUTPUTS: { string - textual result for escape }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: Escape

    private static string Escape(string value)
    {
        // START_BLOCK_ESCAPE_VALUE
        return value.Replace(Separator.ToString(), " ");
        // END_BLOCK_ESCAPE_VALUE
    }

    // START_CONTRACT: NormalizeCellText
    //   PURPOSE: Normalize cell text.
    //   INPUTS: { value: string - method parameter }
    //   OUTPUTS: { string - textual result for normalize cell text }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: NormalizeCellText

    private static string NormalizeCellText(string value)
    {
        // START_BLOCK_NORMALIZE_CELL_TEXT
        return (value ?? string.Empty).Replace('\u00A0', ' ').Trim();
        // END_BLOCK_NORMALIZE_CELL_TEXT
    }

    // START_CONTRACT: NormalizeSingleLineText
    //   PURPOSE: Normalize single line text.
    //   INPUTS: { value: string - method parameter }
    //   OUTPUTS: { string - textual result for normalize single line text }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: NormalizeSingleLineText

    private static string NormalizeSingleLineText(string value)
    {
        // START_BLOCK_NORMALIZE_SINGLE_LINE_TEXT
        string normalized = NormalizeCellText(value)
            .Replace("\r\n", " ")
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Replace('\t', ' ');
        while (normalized.Contains("  ", StringComparison.Ordinal))
        {
            normalized = normalized.Replace("  ", " ", StringComparison.Ordinal);
        }

        return normalized.Trim();
        // END_BLOCK_NORMALIZE_SINGLE_LINE_TEXT
    }

    // START_CONTRACT: NormalizeCableValue
    //   PURPOSE: Normalize cable value.
    //   INPUTS: { raw: string - method parameter }
    //   OUTPUTS: { string - textual result for normalize cable value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: NormalizeCableValue

    private static string NormalizeCableValue(string raw)
    {
        // START_BLOCK_NORMALIZE_CABLE_VALUE
        string singleLine = NormalizeSingleLineText(raw);
        if (string.IsNullOrWhiteSpace(singleLine))
        {
            return string.Empty;
        }

        int lengthTagIndex = singleLine.IndexOf(" L=", StringComparison.OrdinalIgnoreCase);
        if (lengthTagIndex > 0)
        {
            singleLine = singleLine[..lengthTagIndex].Trim();
        }
        else if (singleLine.StartsWith("L=", StringComparison.OrdinalIgnoreCase))
        {
            return string.Empty;
        }

        return singleLine.Trim();
        // END_BLOCK_NORMALIZE_CABLE_VALUE
    }

    // START_CONTRACT: NormalizeForLookup
    //   PURPOSE: Normalize for lookup.
    //   INPUTS: { value: string - method parameter }
    //   OUTPUTS: { string - textual result for normalize for lookup }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: NormalizeForLookup

    private static string NormalizeForLookup(string value)
    {
        // START_BLOCK_NORMALIZE_FOR_LOOKUP
        return NormalizeCellText(value).Replace(" ", string.Empty);
        // END_BLOCK_NORMALIZE_FOR_LOOKUP
    }

    // START_CONTRACT: EnumerateWorkbookNamedRanges
    //   PURPOSE: Enumerate workbook named ranges.
    //   INPUTS: { workbook: XLWorkbook - method parameter }
    //   OUTPUTS: { IEnumerable<NamedRangeWindow> - result of enumerate workbook named ranges }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: EnumerateWorkbookNamedRanges

    private static IEnumerable<NamedRangeWindow> EnumerateWorkbookNamedRanges(XLWorkbook workbook)
    {
        // START_BLOCK_ENUMERATE_WORKBOOK_NAMED_RANGES
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
                object? worksheet = TryGetPropertyValue(range, "Worksheet");
                string worksheetName = ReadStringProperty(worksheet, "Name");
                object? rangeAddress = TryGetPropertyValue(range, "RangeAddress");
                if (rangeAddress is null)
                {
                    continue;
                }

                object? first = TryGetPropertyValue(rangeAddress, "FirstAddress");
                object? last = TryGetPropertyValue(rangeAddress, "LastAddress");
                int minRow = ReadIntProperty(first, "RowNumber");
                int maxRow = ReadIntProperty(last, "RowNumber");
                int minColumn = ReadIntProperty(first, "ColumnNumber");
                int maxColumn = ReadIntProperty(last, "ColumnNumber");
                if (minRow <= 0 || maxRow <= 0 || minColumn <= 0 || maxColumn <= 0)
                {
                    continue;
                }

                string address = rangeAddress.ToString() ?? string.Empty;
                yield return new NamedRangeWindow(name, worksheetName, minRow, maxRow, minColumn, maxColumn, address);
            }
        }
        // END_BLOCK_ENUMERATE_WORKBOOK_NAMED_RANGES
    }

    // START_CONTRACT: EnumerateObjects
    //   PURPOSE: Enumerate objects.
    //   INPUTS: { candidate: object - method parameter }
    //   OUTPUTS: { IEnumerable<object> - result of enumerate objects }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: EnumerateObjects

    private static IEnumerable<object> EnumerateObjects(object candidate)
    {
        // START_BLOCK_ENUMERATE_OBJECTS
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
        // END_BLOCK_ENUMERATE_OBJECTS
    }

    // START_CONTRACT: TryGetPropertyValue
    //   PURPOSE: Attempt to execute get property value.
    //   INPUTS: { instance: object? - method parameter; propertyName: string - method parameter }
    //   OUTPUTS: { object? - result of attempt to execute get property value }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: TryGetPropertyValue

    private static object? TryGetPropertyValue(object? instance, string propertyName)
    {
        // START_BLOCK_TRY_GET_PROPERTY_VALUE
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
        // END_BLOCK_TRY_GET_PROPERTY_VALUE
    }

    // START_CONTRACT: ReadStringProperty
    //   PURPOSE: Read string property.
    //   INPUTS: { instance: object? - method parameter; propertyName: string - method parameter }
    //   OUTPUTS: { string - textual result for read string property }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ReadStringProperty

    private static string ReadStringProperty(object? instance, string propertyName)
    {
        // START_BLOCK_READ_STRING_PROPERTY
        object? value = TryGetPropertyValue(instance, propertyName);
        return value?.ToString()?.Trim() ?? string.Empty;
        // END_BLOCK_READ_STRING_PROPERTY
    }

    // START_CONTRACT: ReadIntProperty
    //   PURPOSE: Read int property.
    //   INPUTS: { instance: object? - method parameter; propertyName: string - method parameter }
    //   OUTPUTS: { int - result of read int property }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-EXCEL-GATEWAY
    // END_CONTRACT: ReadIntProperty

    private static int ReadIntProperty(object? instance, string propertyName)
    {
        // START_BLOCK_READ_INT_PROPERTY
        object? value = TryGetPropertyValue(instance, propertyName);
        if (value is null)
        {
            return 0;
        }

        if (value is int i)
        {
            return i;
        }

        if (value is long l && l >= int.MinValue && l <= int.MaxValue)
        {
            return (int)l;
        }

        if (value is short s)
        {
            return s;
        }

        return int.TryParse(value.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : 0;
        // END_BLOCK_READ_INT_PROPERTY
    }

    private readonly record struct HeaderMap(
        int RowNumber,
        int ShieldColumn,
        int LineColumn,
        int CableColumn,
        int BreakerColumn,
        int NoteColumn)
    {
        public static HeaderMap Empty => new(0, 0, 0, 0, 0, 0);

        // START_CONTRACT: ToDebugString
        //   PURPOSE: To debug string.
        //   INPUTS: none
        //   OUTPUTS: { string - textual result for to debug string }
        //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
        //   LINKS: M-EXCEL-GATEWAY
        // END_CONTRACT: ToDebugString

        public string ToDebugString()
        {
            // START_BLOCK_HEADER_MAP_TO_DEBUG_STRING
            string breaker = BreakerColumn > 0 ? ToColumnName(BreakerColumn) : "-";
            string note = NoteColumn > 0 ? ToColumnName(NoteColumn) : "-";
            return $"header_row={RowNumber}; Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В©Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ={ToColumnName(ShieldColumn)}; Р В Р’В Р вЂ™Р’В Р В Р вЂ Р В РІР‚С™Р РЋРІР‚СњР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р В Р вЂ№Р В Р’В Р В Р РЏ={ToColumnName(LineColumn)}; Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†РІР‚С›РЎС›Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В±Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В»Р В Р’В Р В Р вЂ№Р В Р’В Р В РІР‚В°={ToColumnName(CableColumn)}; Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљРЎС›Р В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р РЋРІвЂћСћ={breaker}; Р В Р’В Р вЂ™Р’В Р В Р Р‹Р РЋРЎСџР В Р’В Р В Р вЂ№Р В Р’В Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В Р Р‹Р вЂ™Р’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’ВµР В Р’В Р В Р вЂ№Р В Р вЂ Р В РІР‚С™Р В Р вЂ№Р В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’В°Р В Р’В Р вЂ™Р’В Р В Р’В Р Р†Р вЂљР’В¦Р В Р’В Р вЂ™Р’В Р В Р Р‹Р Р†Р вЂљР’ВР В Р’В Р вЂ™Р’В Р В РІР‚в„ўР вЂ™Р’Вµ={note}";
            // END_BLOCK_HEADER_MAP_TO_DEBUG_STRING
        }
    }

    
    private readonly record struct ImportWindow(
        int MinRow,
        int MaxRow,
        int MinColumn,
        int MaxColumn,
        string DebugLabel)
    {
        // START_CONTRACT: ContainsColumn
        //   PURPOSE: Contains column.
        //   INPUTS: { column: int - method parameter }
        //   OUTPUTS: { bool - true when method can contains column }
        //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
        //   LINKS: M-EXCEL-GATEWAY
        // END_CONTRACT: ContainsColumn

        public bool ContainsColumn(int column)
        {
            // START_BLOCK_IMPORT_WINDOW_CONTAINS_COLUMN
            return column >= MinColumn && column <= MaxColumn;
            // END_BLOCK_IMPORT_WINDOW_CONTAINS_COLUMN
        }
    }
    private readonly record struct NamedRangeWindow(
        string Name,
        string WorksheetName,
        int MinRow,
        int MaxRow,
        int MinColumn,
        int MaxColumn,
        string Address);
    private readonly record struct WorkbookShadowRow(int RowNumber, IReadOnlyDictionary<string, string> Cells)
    {
        public static WorkbookShadowRow Empty { get; } = new(0, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
    }

    private readonly record struct CellReadResult(bool HasValue, bool IsEmpty, bool HasError, string Value, string ErrorMessage)
    {
        // START_CONTRACT: Empty
        //   PURPOSE: Empty.
        //   INPUTS: { ) - method parameter }
        //   OUTPUTS: { CellReadResult - result of empty }
        //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
        //   LINKS: M-EXCEL-GATEWAY
        // END_CONTRACT: Empty

        public static CellReadResult Empty() => new(false, true, false, string.Empty, string.Empty);
        // START_CONTRACT: FromValue
        //   PURPOSE: From value.
        //   INPUTS: { value): string - method parameter }
        //   OUTPUTS: { CellReadResult - result of from value }
        //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
        //   LINKS: M-EXCEL-GATEWAY
        // END_CONTRACT: FromValue

        public static CellReadResult FromValue(string value) => new(true, false, false, value, string.Empty);
        // START_CONTRACT: Error
        //   PURPOSE: Error.
        //   INPUTS: { error): string - method parameter }
        //   OUTPUTS: { CellReadResult - result of error }
        //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
        //   LINKS: M-EXCEL-GATEWAY
        // END_CONTRACT: Error

        public static CellReadResult Error(string error) => new(false, false, true, string.Empty, error);
    }

    public sealed class WorkbookReadDiagnostics
    {
        private const int MaxValidationIssues = 80;

        public WorkbookReadDiagnostics(string workbookPath)
        {
            WorkbookPath = workbookPath;
        }

        public string WorkbookPath { get; }
        public List<string> SheetNames { get; } = [];
        public string SelectedSheetName { get; set; } = string.Empty;
        public string ImportWindow { get; set; } = string.Empty;
        public int HeaderRowNumber { get; set; }
        public int DataStartRowNumber { get; set; }
        public int DataEndRowNumber { get; set; }
        public string ColumnMap { get; set; } = string.Empty;
        public int AcceptedRows { get; set; }
        public string? FatalError { get; set; }
        public List<string> PreviewRows { get; } = [];
        public List<string> ValidationIssues { get; } = [];
        public int SuppressedValidationIssueCount { get; private set; }

        // START_CONTRACT: AddValidationIssue
        //   PURPOSE: Add validation issue.
        //   INPUTS: { issue: string - method parameter }
        //   OUTPUTS: { void - no return value }
        //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
        //   LINKS: M-EXCEL-GATEWAY
        // END_CONTRACT: AddValidationIssue

        public void AddValidationIssue(string issue)
        {
            // START_BLOCK_ADD_VALIDATION_ISSUE
            if (string.IsNullOrWhiteSpace(issue))
            {
                return;
            }

            if (ValidationIssues.Count < MaxValidationIssues)
            {
                ValidationIssues.Add(issue);
            }
            else
            {
                SuppressedValidationIssueCount++;
            }
            // END_BLOCK_ADD_VALIDATION_ISSUE
        }
    }
}