// FILE: src/Commands/CommandRegistry.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Register AutoCAD commands and dispatch to domain services.
//   SCOPE: EOM_MAP, EOM_TRACE, EOM_SPEC command handlers.
//   DEPENDS: M-MAP, M-TRACE, M-SPEC, M-EXPORT, M-LICENSE, M-LOG
//   LINKS: M-COMMANDS, M-MAP, M-TRACE, M-SPEC, M-EXPORT, M-LICENSE, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   EomMap - Executes block mapping workflow.
//   EomTrace - Executes cable trace workflow.
//   EomSpec - Executes specification workflow.
//   EomMapCfg - Opens mapping configuration window.
// END_MODULE_MAP
//
// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.1.0 - Updated EOM_MAP to direct source/target selection and replacement without attribute transfer.
// END_CHANGE_SUMMARY

using Autodesk.AutoCAD.Runtime;
using ElTools.Data;
using ElTools.Integrations;
using ElTools.Models;
using ElTools.Services;
using ElTools.Shared;
using ElTools.UI;

namespace ElTools.Commands;

public class CommandRegistry
{
    private readonly BlockMappingService _mapping = new();
    private readonly CableTraceService _trace = new();
    private readonly SpecificationService _spec = new();
    private readonly SpecExportService _export = new();
    private readonly LicenseService _license = new();
    private readonly LogService _log = new();
    private readonly SettingsRepository _settings = new();
    private readonly XDataService _xdata = new();
    private readonly AutoCADAdapter _acad = new();

    public CommandRegistry()
    {
        _acad.SubscribeObjectAppended(OnObjectAppended);
    }

    [CommandMethod("EOM_MAP")]
    public void EomMap()
    {
        // START_BLOCK_COMMAND_EOM_MAP
        if (_license.Validate() is LicenseState.Invalid or LicenseState.Expired)
        {
            _log.Write("Команда EOM_MAP заблокирована лицензией.");
            return;
        }

        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            _log.Write("Активный документ не найден.");
            return;
        }

        Editor editor = doc.Editor;
        var sourceOptions = new PromptEntityOptions("\nУкажите исходный блок для замены: ");
        sourceOptions.SetRejectMessage("\nНужен блок (BlockReference).");
        sourceOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult sourceResult = editor.GetEntity(sourceOptions);
        if (sourceResult.Status != PromptStatus.OK)
        {
            _log.Write("Команда отменена: исходный блок не выбран.");
            return;
        }

        var targetOptions = new PromptEntityOptions("\nУкажите целевой блок: ");
        targetOptions.SetRejectMessage("\nНужен блок (BlockReference).");
        targetOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult targetResult = editor.GetEntity(targetOptions);
        if (targetResult.Status != PromptStatus.OK)
        {
            _log.Write("Команда отменена: целевой блок не выбран.");
            return;
        }

        int replaced = _mapping.ExecuteMapping(sourceResult.ObjectId, targetResult.ObjectId);
        _log.Write($"EOM_MAP завершена. Заменено блоков: {replaced}.");
        // END_BLOCK_COMMAND_EOM_MAP
    }

    [CommandMethod("EOM_TRACE")]
    public void EomTrace()
    {
        // START_BLOCK_COMMAND_EOM_TRACE
        if (_license.Validate() is LicenseState.Invalid or LicenseState.Expired)
        {
            _log.Write("Команда EOM_TRACE заблокирована лицензией.");
            return;
        }

        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        Editor editor = doc.Editor;

        var baseLineOptions = new PromptEntityOptions("\nУкажите базовую полилинию (магистраль): ");
        baseLineOptions.SetRejectMessage("\nНужна полилиния.");
        baseLineOptions.AddAllowedClass(typeof(Polyline), true);
        PromptEntityResult baseLineResult = editor.GetEntity(baseLineOptions);
        if (baseLineResult.Status != PromptStatus.OK)
        {
            _log.Write("Команда отменена: базовая полилиния не выбрана.");
            return;
        }

        int created = 0;
        while (true)
        {
            var targetBlockOptions = new PromptEntityOptions("\nУкажите блок для ответвления [Enter - завершить]: ")
            {
                AllowNone = true
            };
            targetBlockOptions.SetRejectMessage("\nНужен блок (BlockReference).");
            targetBlockOptions.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult targetBlockResult = editor.GetEntity(targetBlockOptions);

            if (targetBlockResult.Status == PromptStatus.None)
            {
                _log.Write($"EOM_TRACE завершена. Построено ответвлений: {created}.");
                break;
            }

            if (targetBlockResult.Status != PromptStatus.OK)
            {
                _log.Write("Команда отменена пользователем.");
                break;
            }

            TraceResult? trace = _trace.ExecuteTraceFromBase(baseLineResult.ObjectId, targetBlockResult.ObjectId);
            if (trace is null)
            {
                _log.Write("Не удалось построить ответвление для выбранного блока.");
                continue;
            }

            created++;
            _log.Write($"Ответвление #{created}: {trace.TotalLength:0.###} м.");
        }
        // END_BLOCK_COMMAND_EOM_TRACE
    }

    [CommandMethod("EOM_SPEC")]
    public void EomSpec()
    {
        // START_BLOCK_COMMAND_EOM_SPEC
        if (_license.Validate() is LicenseState.Invalid or LicenseState.Expired)
        {
            _log.Write("Команда EOM_SPEC заблокирована лицензией.");
            return;
        }

        IReadOnlyList<SpecificationRow> rows = _spec.BuildSpecification(Array.Empty<ObjectId>());
        _export.ToAutoCadTable(rows);
        _export.ToCsv(rows);
        _log.Write("EOM_SPEC завершена.");
        // END_BLOCK_COMMAND_EOM_SPEC
    }

    [CommandMethod("EOM_MAPCFG")]
    public void EomMapCfg()
    {
        // START_BLOCK_COMMAND_EOM_MAPCFG
        try
        {
            var vm = new MappingConfigWindowViewModel();
            var window = new MappingConfigWindow(vm);
            Application.ShowModalWindow(window);
            _log.Write("Окно настройки соответствий закрыто.");
        }
        catch (System.Exception ex)
        {
            _log.Write($"Ошибка открытия окна настройки соответствий: {ex.Message}");
        }
        // END_BLOCK_COMMAND_EOM_MAPCFG
    }

    [CommandMethod(PluginConfig.Commands.ActiveGroup)]
    public void EomActiveGroup()
    {
        // START_BLOCK_COMMAND_EOM_ACTIVE_GROUP
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PromptResult prompt = doc.Editor.GetString("\nВведите код активной группы: ");
        if (prompt.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(prompt.StringResult))
        {
            _log.Write("Активная группа не установлена.");
            return;
        }

        _xdata.SetActiveGroup(prompt.StringResult.Trim());
        _acad.SubscribeObjectAppended(OnObjectAppended);
        _log.Write($"Активная группа установлена: {prompt.StringResult.Trim()}");
        // END_BLOCK_COMMAND_EOM_ACTIVE_GROUP
    }

    [CommandMethod(PluginConfig.Commands.AssignGroup)]
    public void EomAssignGroup()
    {
        // START_BLOCK_COMMAND_EOM_ASSIGN_GROUP
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PromptResult groupPrompt = doc.Editor.GetString("\nВведите код группы: ");
        if (groupPrompt.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(groupPrompt.StringResult))
        {
            _log.Write("Группа не задана.");
            return;
        }

        var selectionOptions = new PromptSelectionOptions { MessageForAdding = "\nВыберите линии/полилинии: " };
        var filter = new SelectionFilter([
            new TypedValue((int)DxfCode.Start, "LINE,LWPOLYLINE")
        ]);
        PromptSelectionResult selection = doc.Editor.GetSelection(selectionOptions, filter);
        if (selection.Status != PromptStatus.OK)
        {
            _log.Write("Назначение группы отменено.");
            return;
        }

        _xdata.AssignGroupToSelection(selection.Value.GetObjectIds(), groupPrompt.StringResult.Trim());
        _log.Write($"Группа {groupPrompt.StringResult.Trim()} назначена: {selection.Value.Count} объектов.");
        // END_BLOCK_COMMAND_EOM_ASSIGN_GROUP
    }

    [CommandMethod(PluginConfig.Commands.InstallTypeSettings)]
    public void EomInstallTypeSettings()
    {
        // START_BLOCK_COMMAND_EOM_INSTALL_TYPE_SETTINGS
        string path = _settings.OpenInstallTypeConfig();
        _log.Write($"Конфиг правил прокладки: {path}");
        // END_BLOCK_COMMAND_EOM_INSTALL_TYPE_SETTINGS
    }

    [CommandMethod(PluginConfig.Commands.Update)]
    public void EomUpdate()
    {
        // START_BLOCK_COMMAND_EOM_UPDATE
        EomTrace();
        _log.Write("EOM_ОБНОВИТЬ завершена.");
        // END_BLOCK_COMMAND_EOM_UPDATE
    }

    [CommandMethod(PluginConfig.Commands.ExportExcel)]
    public void EomExportExcel()
    {
        // START_BLOCK_COMMAND_EOM_EXPORT_EXCEL
        IReadOnlyList<SpecificationRow> specRows = _spec.BuildSpecification(Array.Empty<ObjectId>());
        string templatePath = _settings.LoadSettings().ExcelTemplatePath;
        var inputRows = specRows.Select(ToExcelInputRow).ToList();
        _export.ToExcelInput(templatePath, inputRows);
        _log.Write("EOM_ЭКСПОРТ_EXCEL завершена.");
        // END_BLOCK_COMMAND_EOM_EXPORT_EXCEL
    }

    [CommandMethod(PluginConfig.Commands.ImportExcel)]
    public void EomImportExcel()
    {
        // START_BLOCK_COMMAND_EOM_IMPORT_EXCEL
        string templatePath = _settings.LoadSettings().ExcelTemplatePath;
        IReadOnlyList<ExcelOutputRow> rows = _export.FromExcelOutput(templatePath);
        _log.Write($"EOM_ИМПОРТ_EXCEL завершена. Импортировано строк: {rows.Count}.");
        // END_BLOCK_COMMAND_EOM_IMPORT_EXCEL
    }

    [CommandMethod(PluginConfig.Commands.BuildOls)]
    public void EomBuildOls()
    {
        // START_BLOCK_COMMAND_EOM_BUILD_OLS
        string templatePath = _settings.LoadSettings().ExcelTemplatePath;
        IReadOnlyList<ExcelOutputRow> rows = _export.FromExcelOutput(templatePath);
        _log.Write($"EOM_ПОСТРОИТЬ_ОЛС: подготовлено строк для построения: {rows.Count}.");
        // END_BLOCK_COMMAND_EOM_BUILD_OLS
    }

    [CommandMethod(PluginConfig.Commands.BuildPanelLayout)]
    public void EomBuildPanelLayout()
    {
        // START_BLOCK_COMMAND_EOM_BUILD_PANEL_LAYOUT
        string templatePath = _settings.LoadSettings().ExcelTemplatePath;
        IReadOnlyList<ExcelOutputRow> rows = _export.FromExcelOutput(templatePath);
        int modulesPerRow = _settings.LoadSettings().PanelModulesPerRow;
        _log.Write($"EOM_КОМПОНОВКА_ЩИТА: строк {rows.Count}, модулей в ряду {modulesPerRow}.");
        // END_BLOCK_COMMAND_EOM_BUILD_PANEL_LAYOUT
    }

    [CommandMethod(PluginConfig.Commands.Validate)]
    public void EomValidate()
    {
        // START_BLOCK_COMMAND_EOM_VALIDATE
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        int missingGroupLines = 0;
        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            var blockTable = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in modelSpace)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is not Entity entity || entity is not (Line or Polyline))
                {
                    continue;
                }

                string? group = _xdata.GetLineGroup(id);
                if (string.IsNullOrWhiteSpace(group))
                {
                    missingGroupLines++;
                }
            }

            tr.Commit();
        }

        _log.Write($"EOM_ПРОВЕРКА: линий без ГРУППА: {missingGroupLines}.");
        // END_BLOCK_COMMAND_EOM_VALIDATE
    }

    private void OnObjectAppended(ObjectId objectId)
    {
        // START_BLOCK_ON_OBJECT_APPENDED
        _xdata.ApplyActiveGroupToEntity(objectId);
        // END_BLOCK_ON_OBJECT_APPENDED
    }

    private static ExcelInputRow ToExcelInputRow(SpecificationRow row)
    {
        // START_BLOCK_TO_EXCEL_INPUT_ROW
        string group = row.Group;
        double total = row.TotalLength;
        double ceiling = row.CableType.Equals("Потолок", StringComparison.OrdinalIgnoreCase) ? total : 0;
        double floor = row.CableType.Equals("Пол", StringComparison.OrdinalIgnoreCase) ? total : 0;
        double riser = row.CableType.Equals("Стояк", StringComparison.OrdinalIgnoreCase) ? total : 0;

        return new ExcelInputRow(
            Shield: string.Empty,
            Group: group,
            PowerKw: 0,
            Voltage: 0,
            TotalLengthMeters: total,
            CeilingLengthMeters: ceiling,
            FloorLengthMeters: floor,
            RiserLengthMeters: riser);
        // END_BLOCK_TO_EXCEL_INPUT_ROW
    }
}
