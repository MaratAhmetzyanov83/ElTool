// FILE: src/Commands/CommandRegistry.cs
// VERSION: 1.6.0
// START_MODULE_CONTRACT
//   PURPOSE: Register and execute AutoCAD entry commands for mapping, tracing, Excel workflows, OLS generation, panel layout, and validation.
//   SCOPE: Full command lifecycle from user prompts to orchestrator/service calls and drawing output generation.
//   DEPENDS: M-MAP-ORCHESTRATOR, M-TRACE-ORCHESTRATOR, M-SPEC-ORCHESTRATOR, M-SELECTION, M-EXPORT, M-SETTINGS, M-XDATA, M-CAD-CONTEXT, M-LOGGING, M-LICENSE, M-PANEL-LAYOUT-CONFIG-UI, M-PANEL-LAYOUT-CONFIG-VM
//   LINKS: M-ENTRY-COMMANDS, M-MAP-ORCHESTRATOR, M-TRACE-ORCHESTRATOR, M-SPEC-ORCHESTRATOR, M-SELECTION, M-EXPORT, M-SETTINGS, M-XDATA, M-CAD-CONTEXT, M-LOGGING, M-CONFIG, M-PANEL-LAYOUT-CONFIG-UI, M-PANEL-LAYOUT-CONFIG-VM
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   EomMap - Executes block mapping workflow.
//   EomTrace - Executes cable trace workflow.
//   EomSpec - Executes specification workflow.
//   EomBuildOls - Draws one-line diagram from Excel OUTPUT.
//   EomPanelLayoutConfig - Opens UI editor for panel layout map and SOURCE->LAYOUT bindings.
//   EomBuildPanelLayout - Builds panel layout from user-selected OLS blocks using panel mapping configuration.
//   EomBindPanelLayoutVisualization - Creates and stores source OLS to visualization block selector rule.
//   BuildPanelLayoutModel - Converts mapped OLS devices to normalized DIN placement model.
// END_MODULE_MAP
//
// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.6.0 - Added panel layout configuration UI with drawing-based SOURCE/LAYOUT binding selection.
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
    private readonly IInstallTypeResolver _installTypeResolver = new InstallTypeResolver();

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

        if (!TryResolveExcelTemplatePath(PluginConfig.Commands.Spec, out string templatePath))
        {
            return;
        }

        IReadOnlyList<SpecificationRow> rows = _spec.BuildSpecification(Array.Empty<ObjectId>());
        _export.ToAutoCadTable(rows);
        _export.ToCsv(rows);
        IReadOnlyList<GroupTraceAggregate> aggregates = _trace.RecalculateByGroups();
        _export.ToExcelInput(templatePath, aggregates.Select(ToExcelInputRow).ToList());
        bool cacheHit = _export.TryGetCachedOutput(templatePath, out IReadOnlyList<ExcelOutputRow> outputRows);
        if (!cacheHit)
        {
            string outputPath = GetExpectedExcelOutputPath(templatePath);
            if (File.Exists(outputPath))
            {
                outputRows = _export.GetCachedOrLoadOutput(templatePath);
            }
            else
            {
                _log.Write($"EOM_SPEC: файл OUTPUT не найден ({outputPath}). Выполните расчет и запустите EOM_ИМПОРТ_EXCEL.");
            }
        }

        if (outputRows.Count > 0)
        {
            _export.ExportExcelOutputReportCsv(outputRows);
            _export.ToAutoCadTableFromOutput(outputRows, new Point3d(0, -120, 0));
        }

        string inputPath = GetExpectedExcelInputPath(templatePath);
        _log.Write($"EOM_SPEC завершена. INPUT: {inputPath}");
        // END_BLOCK_COMMAND_EOM_SPEC
    }

    [CommandMethod(PluginConfig.Commands.MapConfig)]
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

    [CommandMethod(PluginConfig.Commands.PanelLayoutConfig)]
    public void EomPanelLayoutConfig()
    {
        // START_BLOCK_COMMAND_EOM_PANEL_LAYOUT_CONFIG
        try
        {
            var vm = new PanelLayoutConfigWindowViewModel(
                PickPanelLayoutSourceSignatureFromDrawing,
                PickPanelLayoutVisualizationBlockFromDrawing);
            var window = new PanelLayoutConfigWindow(vm);
            Application.ShowModalWindow(window);
            _log.Write("Окно настройки компоновки щита закрыто.");
        }
        catch (System.Exception ex)
        {
            _log.Write($"Ошибка открытия окна настройки компоновки щита: {ex.Message}");
        }
        // END_BLOCK_COMMAND_EOM_PANEL_LAYOUT_CONFIG
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
        IReadOnlyList<GroupTraceAggregate> aggregates = _trace.RecalculateByGroups();
        _log.Write($"EOM_ОБНОВИТЬ завершена. Групп в расчете: {aggregates.Count}.");
        // END_BLOCK_COMMAND_EOM_UPDATE
    }

    [CommandMethod(PluginConfig.Commands.ExportExcel)]
    public void EomExportExcel()
    {
        // START_BLOCK_COMMAND_EOM_EXPORT_EXCEL
        if (!TryResolveExcelTemplatePath(PluginConfig.Commands.ExportExcel, out string templatePath))
        {
            return;
        }

        IReadOnlyList<GroupTraceAggregate> aggregates = _trace.RecalculateByGroups();
        _export.ClearCachedOutput(templatePath);
        var inputRows = aggregates.Select(ToExcelInputRow).ToList();
        _export.ToExcelInput(templatePath, inputRows);
        _log.Write($"EOM_ЭКСПОРТ_EXCEL завершена. INPUT: {GetExpectedExcelInputPath(templatePath)}");
        // END_BLOCK_COMMAND_EOM_EXPORT_EXCEL
    }

    [CommandMethod(PluginConfig.Commands.ImportExcel)]
    public void EomImportExcel()
    {
        // START_BLOCK_COMMAND_EOM_IMPORT_EXCEL
        if (!TryResolveExcelTemplatePath(PluginConfig.Commands.ImportExcel, out string templatePath))
        {
            return;
        }

        string outputPath = GetExpectedExcelOutputPath(templatePath);
        if (!File.Exists(outputPath))
        {
            _log.Write($"EOM_ИМПОРТ_EXCEL: не найден файл OUTPUT: {outputPath}. Сначала заполните расчет и сохраните OUTPUT.");
            return;
        }

        IReadOnlyList<ExcelOutputRow> rows = _export.FromExcelOutput(templatePath);
        if (rows.Count > 0)
        {
            _export.ExportExcelOutputReportCsv(rows);
        }

        _log.Write($"EOM_ИМПОРТ_EXCEL завершена. Импортировано строк: {rows.Count}. Кэш обновлен.");
        // END_BLOCK_COMMAND_EOM_IMPORT_EXCEL
    }

    [CommandMethod(PluginConfig.Commands.BuildOls)]
    public void EomBuildOls()
    {
        // START_BLOCK_COMMAND_EOM_BUILD_OLS
        if (!TryResolveExcelTemplatePath(PluginConfig.Commands.BuildOls, out string templatePath))
        {
            return;
        }

        string outputPath = GetExpectedExcelOutputPath(templatePath);
        bool cacheHit = _export.TryGetCachedOutput(templatePath, out IReadOnlyList<ExcelOutputRow> rows);
        if (!cacheHit)
        {
            if (!File.Exists(outputPath))
            {
                _log.Write($"EOM_ПОСТРОИТЬ_ОЛС: не найден OUTPUT ({outputPath}). Выполните EOM_ИМПОРТ_EXCEL.");
                return;
            }

            rows = _export.GetCachedOrLoadOutput(templatePath);
        }

        if (rows.Count == 0)
        {
            _log.Write($"EOM_ПОСТРОИТЬ_ОЛС: OUTPUT пустой ({outputPath}).");
            return;
        }

        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PromptPointResult pointResult = doc.Editor.GetPoint("\nУкажите точку вставки ОЛС: ");
        if (pointResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_ПОСТРОИТЬ_ОЛС отменена.");
            return;
        }

        PromptResult shieldResult = doc.Editor.GetString("\nЩИТ для построения [Enter = все]: ");
        string? shield = shieldResult.Status == PromptStatus.OK ? shieldResult.StringResult?.Trim() : null;
        IReadOnlyList<ExcelOutputRow> sourceRows = string.IsNullOrWhiteSpace(shield)
            ? rows
            : rows.Where(x => string.Equals(x.Shield, shield, StringComparison.OrdinalIgnoreCase)).ToList();

        DrawOlsRows(sourceRows, pointResult.Value);
        _log.Write($"EOM_ПОСТРОИТЬ_ОЛС: построено строк: {sourceRows.Count}. Источник={(cacheHit ? "кэш" : "файл")}.");
        // END_BLOCK_COMMAND_EOM_BUILD_OLS
    }

    [CommandMethod(PluginConfig.Commands.BuildPanelLayout)]
    public void EomBuildPanelLayout()
    {
        // START_BLOCK_COMMAND_EOM_BUILD_PANEL_LAYOUT
        SettingsModel settings = _settings.LoadSettings();
        PanelLayoutMapConfig mapConfig = _settings.LoadPanelLayoutMap();
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PromptSelectionOptions selectionOptions = new()
        {
            MessageForAdding = "\nВыделите готовую однолинейную схему (блоки аппаратов): "
        };
        SelectionFilter filter = new([
            new TypedValue((int)DxfCode.Start, "INSERT")
        ]);
        PromptSelectionResult selection = doc.Editor.GetSelection(selectionOptions, filter);
        if (selection.Status != PromptStatus.OK)
        {
            _log.Write("EOM_КОМПОНОВКА_ЩИТА отменена: ОЛС не выделена.");
            return;
        }

        ObjectId[] selectedIds = selection.Value.GetObjectIds();
        if (selectedIds.Length == 0)
        {
            _log.Write("EOM_КОМПОНОВКА_ЩИТА: не выбраны блоки ОЛС.");
            return;
        }

        ParsedOlsSelection parsedSelection = ParseOlsSelection(doc, selectedIds, mapConfig.AttributeTags);
        IReadOnlyList<MappedLayoutDevice> mappedDevices = MapOlsDevices(
            parsedSelection.Devices,
            mapConfig.SelectorRules ?? [],
            mapConfig.LayoutMap ?? [],
            out List<SkippedOlsDeviceIssue> mappingIssues);
        var issues = new List<SkippedOlsDeviceIssue>(parsedSelection.Issues.Count + mappingIssues.Count);
        issues.AddRange(parsedSelection.Issues);
        issues.AddRange(mappingIssues);

        if (mappedDevices.Count == 0)
        {
            ReportSkippedOlsDevices(issues);
            _log.Write("EOM_КОМПОНОВКА_ЩИТА: нет валидных правил сопоставления (нужны SelectorRules либо legacy LayoutMap + МОДУЛЕЙ/FallbackModules).");
            return;
        }

        int defaultModulesPerRow = mapConfig.DefaultModulesPerRow > 0
            ? mapConfig.DefaultModulesPerRow
            : (settings.PanelModulesPerRow > 0 ? settings.PanelModulesPerRow : 24);
        PromptIntegerOptions modulesOptions = new($"\nМодулей в ряду [Enter = {defaultModulesPerRow}]: ")
        {
            AllowNone = true,
            LowerLimit = 1,
            UpperLimit = 72
        };
        PromptIntegerResult modulesResult = doc.Editor.GetInteger(modulesOptions);
        if (modulesResult.Status == PromptStatus.Cancel)
        {
            _log.Write("EOM_КОМПОНОВКА_ЩИТА отменена.");
            return;
        }

        int modulesPerRow = modulesResult.Status == PromptStatus.OK ? modulesResult.Value : defaultModulesPerRow;
        PromptPointResult pointResult = doc.Editor.GetPoint("\nУкажите точку вставки компоновки щита: ");
        if (pointResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_КОМПОНОВКА_ЩИТА отменена.");
            return;
        }

        IReadOnlyList<PanelLayoutRow> layoutRows = BuildPanelLayoutModel(mappedDevices, modulesPerRow);
        if (layoutRows.Count == 0)
        {
            _log.Write("EOM_КОМПОНОВКА_ЩИТА: не сформирована модель раскладки.");
            return;
        }

        int renderedSegments = DrawPanelLayout(layoutRows, pointResult.Value, modulesPerRow, out List<SkippedOlsDeviceIssue> renderIssues);
        issues.AddRange(renderIssues);
        ReportSkippedOlsDevices(issues);

        int dinRows = layoutRows.Max(x => x.DinRow);
        int uniqueDevices = layoutRows
            .Select(x => x.EntityId)
            .Distinct()
            .Count();
        _log.Write(
            $"EOM_КОМПОНОВКА_ЩИТА: выделено блоков {selectedIds.Length}, валидных блоков {parsedSelection.Devices.Count}, отрисовано устройств {uniqueDevices}, сегментов {renderedSegments}, рядов DIN {dinRows}, пропущено {issues.Count}.");
        // END_BLOCK_COMMAND_EOM_BUILD_PANEL_LAYOUT
    }

    [CommandMethod(PluginConfig.Commands.BindPanelLayoutVisualization)]
    public void EomBindPanelLayoutVisualization()
    {
        // START_BLOCK_COMMAND_EOM_BIND_PANEL_LAYOUT_VISUALIZATION
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PanelLayoutMapConfig mapConfig = _settings.LoadPanelLayoutMap();

        var sourceOptions = new PromptEntityOptions("\nВыберите исходный блок ОЛС для привязки: ");
        sourceOptions.SetRejectMessage("\nНужен блок (BlockReference).");
        sourceOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult sourceResult = doc.Editor.GetEntity(sourceOptions);
        if (sourceResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_СВЯЗАТЬ_ВИЗУАЛИЗАЦИЮ отменена: исходный блок ОЛС не выбран.");
            return;
        }

        if (!TryReadSourceSignature(doc, sourceResult.ObjectId, out OlsSourceSignature signature))
        {
            _log.Write("EOM_СВЯЗАТЬ_ВИЗУАЛИЗАЦИЮ: не удалось прочитать сигнатуру исходного блока.");
            return;
        }

        var targetOptions = new PromptEntityOptions("\nВыберите блок визуализации для компоновки щита: ");
        targetOptions.SetRejectMessage("\nНужен блок (BlockReference).");
        targetOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult targetResult = doc.Editor.GetEntity(targetOptions);
        if (targetResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_СВЯЗАТЬ_ВИЗУАЛИЗАЦИЮ отменена: блок визуализации не выбран.");
            return;
        }

        if (!TryReadEffectiveBlockName(doc, targetResult.ObjectId, out string layoutBlockName))
        {
            _log.Write("EOM_СВЯЗАТЬ_ВИЗУАЛИЗАЦИЮ: не удалось определить имя блока визуализации.");
            return;
        }

        PromptIntegerOptions priorityOptions = new("\nПриоритет правила [Enter = 100]: ")
        {
            AllowNone = true,
            LowerLimit = 0,
            UpperLimit = 10000
        };
        PromptIntegerResult priorityResult = doc.Editor.GetInteger(priorityOptions);
        if (priorityResult.Status == PromptStatus.Cancel)
        {
            _log.Write("EOM_СВЯЗАТЬ_ВИЗУАЛИЗАЦИЮ отменена.");
            return;
        }

        int priority = priorityResult.Status == PromptStatus.OK ? priorityResult.Value : 100;
        PromptIntegerOptions fallbackOptions = new("\nFallback модулей [Enter = без fallback]: ")
        {
            AllowNone = true,
            LowerLimit = 1,
            UpperLimit = 72
        };
        PromptIntegerResult fallbackResult = doc.Editor.GetInteger(fallbackOptions);
        if (fallbackResult.Status == PromptStatus.Cancel)
        {
            _log.Write("EOM_СВЯЗАТЬ_ВИЗУАЛИЗАЦИЮ отменена.");
            return;
        }

        int? fallbackModules = fallbackResult.Status == PromptStatus.OK ? fallbackResult.Value : null;
        string normalizedSourceName = signature.SourceBlockName.Trim();
        string? normalizedVisibility = string.IsNullOrWhiteSpace(signature.VisibilityValue) ? null : signature.VisibilityValue.Trim();
        string normalizedLayoutName = layoutBlockName.Trim();

        PanelLayoutSelectorRule newRule = new(
            Priority: priority,
            SourceBlockName: normalizedSourceName,
            VisibilityValue: normalizedVisibility,
            LayoutBlockName: normalizedLayoutName,
            FallbackModules: fallbackModules);

        List<PanelLayoutSelectorRule> existingRules = (mapConfig.SelectorRules ?? [])
            .Where(x => x is not null)
            .ToList();

        int existingIndex = existingRules.FindIndex(x =>
            string.Equals(x.SourceBlockName, normalizedSourceName, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.VisibilityValue ?? string.Empty, normalizedVisibility ?? string.Empty, StringComparison.OrdinalIgnoreCase));

        if (existingIndex >= 0)
        {
            existingRules[existingIndex] = newRule;
        }
        else
        {
            existingRules.Add(newRule);
        }

        mapConfig.SelectorRules = existingRules
            .OrderBy(x => x.Priority)
            .ThenBy(x => x.SourceBlockName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.VisibilityValue ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ToList();

        _settings.SavePanelLayoutMap(mapConfig);

        string visibilityPart = string.IsNullOrWhiteSpace(normalizedVisibility)
            ? "*"
            : normalizedVisibility;
        string fallbackPart = fallbackModules.HasValue
            ? fallbackModules.Value.ToString()
            : "none";
        _log.Write(
            $"EOM_СВЯЗАТЬ_ВИЗУАЛИЗАЦИЮ: сохранено правило SOURCE='{normalizedSourceName}', VISIBILITY='{visibilityPart}' -> LAYOUT='{normalizedLayoutName}', PRIORITY={priority}, FALLBACK_MODULES={fallbackPart}.");
        // END_BLOCK_COMMAND_EOM_BIND_PANEL_LAYOUT_VISUALIZATION
    }

    [CommandMethod(PluginConfig.Commands.BindPanelLayoutVisualizationAlias)]
    public void EomBindPanelLayoutVisualizationAlias()
    {
        // START_BLOCK_COMMAND_EOM_BIND_PANEL_LAYOUT_VISUALIZATION_ALIAS
        EomBindPanelLayoutVisualization();
        // END_BLOCK_COMMAND_EOM_BIND_PANEL_LAYOUT_VISUALIZATION_ALIAS
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
        int missingGroupLoads = 0;
        int invalidPowerLoads = 0;
        int groupShieldMismatch = 0;
        int groupsWithoutLines = 0;
        int defaultInstallTypeLines = 0;
        int regexMismatchGroups = 0;
        SettingsModel settings = _settings.LoadSettings();
        InstallTypeRuleSet rules = _settings.LoadInstallTypeRules();
        var groupRegex = string.IsNullOrWhiteSpace(settings.GroupRegex)
            ? null
            : new System.Text.RegularExpressions.Regex(settings.GroupRegex, System.Text.RegularExpressions.RegexOptions.Compiled);

        var lineGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var loadGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var shieldsByGroup = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var loadIdsByGroup = new Dictionary<string, List<ObjectId>>(StringComparer.OrdinalIgnoreCase);
        var issues = new List<ValidationIssue>();
        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            var blockTable = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in modelSpace)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is not Entity entity)
                {
                    continue;
                }

                if (entity is Line or Polyline)
                {
                    string? lineGroup = ReadLineGroup(entity);
                    if (string.IsNullOrWhiteSpace(lineGroup))
                    {
                        missingGroupLines++;
                        issues.Add(new ValidationIssue("LINE_NO_GROUP", id, "Линия без ГРУППА", "Error"));
                    }
                    else
                    {
                        string groupValue = lineGroup.Trim();
                        lineGroups.Add(groupValue);
                        string linetype = _acad.ResolveLinetype(entity);
                        string installType = _installTypeResolver.Resolve(linetype, entity.Layer, rules);
                        if (string.Equals(installType, rules.Default, StringComparison.OrdinalIgnoreCase))
                        {
                            defaultInstallTypeLines++;
                            issues.Add(new ValidationIssue("LINE_DEFAULT_INSTALL_TYPE", id, "Линия с типом прокладки по умолчанию", "Warning"));
                        }
                    }
                }

                if (entity is BlockReference)
                {
                    IReadOnlyDictionary<string, string> attrs = ReadBlockAttributes(tr, (BlockReference)entity);
                    if (attrs.Count == 0)
                    {
                        continue;
                    }

                    attrs.TryGetValue(PluginConfig.AttributeTags.Group, out string? group);
                    attrs.TryGetValue(PluginConfig.AttributeTags.Power, out string? power);
                    attrs.TryGetValue(PluginConfig.AttributeTags.Shield, out string? shield);

                    if (string.IsNullOrWhiteSpace(group))
                    {
                        missingGroupLoads++;
                        issues.Add(new ValidationIssue("LOAD_NO_GROUP", id, "Нагрузка без ГРУППА", "Error"));
                        continue;
                    }

                    string groupValue = group.Trim();
                    loadGroups.Add(groupValue);
                    if (groupRegex is not null && !groupRegex.IsMatch(groupValue))
                    {
                        regexMismatchGroups++;
                        issues.Add(new ValidationIssue("GROUP_REGEX_MISMATCH", id, $"Нестандартный формат ГРУППА: {groupValue}", "Warning"));
                    }

                    if (string.IsNullOrWhiteSpace(power) || !double.TryParse(power.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
                    {
                        invalidPowerLoads++;
                        issues.Add(new ValidationIssue("LOAD_POWER_INVALID", id, "МОЩНОСТЬ отсутствует или нечисловая", "Error"));
                    }

                    if (!shieldsByGroup.TryGetValue(groupValue, out HashSet<string>? shields))
                    {
                        shields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        shieldsByGroup[groupValue] = shields;
                    }
                    if (!loadIdsByGroup.TryGetValue(groupValue, out List<ObjectId>? groupIds))
                    {
                        groupIds = new List<ObjectId>();
                        loadIdsByGroup[groupValue] = groupIds;
                    }
                    groupIds.Add(id);

                    if (!string.IsNullOrWhiteSpace(shield))
                    {
                        shields.Add(shield.Trim());
                    }
                }
            }

            tr.Commit();
        }

        groupShieldMismatch = shieldsByGroup.Values.Count(x => x.Count > 1);
        groupsWithoutLines = loadGroups.Count(x => !lineGroups.Contains(x));
        foreach (string group in loadGroups.Where(x => !lineGroups.Contains(x)))
        {
            if (!loadIdsByGroup.TryGetValue(group, out List<ObjectId>? ids))
            {
                continue;
            }

            foreach (ObjectId id in ids)
            {
                issues.Add(new ValidationIssue("GROUP_WITHOUT_LINES", id, $"Группа {group} есть у нагрузок, но нет линий", "Warning"));
            }
        }

        foreach ((string group, HashSet<string> shields) in shieldsByGroup.Where(x => x.Value.Count > 1))
        {
            if (!loadIdsByGroup.TryGetValue(group, out List<ObjectId>? ids))
            {
                continue;
            }

            foreach (ObjectId id in ids)
            {
                issues.Add(new ValidationIssue("GROUP_SHIELD_MISMATCH", id, $"Группа {group} содержит несколько щитов: {string.Join(", ", shields)}", "Warning"));
            }
        }

        _log.Write(
            $"EOM_ПРОВЕРКА: линии без {PluginConfig.Strings.Group}={missingGroupLines}; " +
            $"нагрузки без {PluginConfig.Strings.Group}={missingGroupLoads}; " +
            $"{PluginConfig.Strings.Power} нечисловая={invalidPowerLoads}; " +
            $"группы нагрузок без линий={groupsWithoutLines}; " +
            $"линии с типом по умолчанию={defaultInstallTypeLines}; " +
            $"несовпадение {PluginConfig.Strings.Shield} внутри группы={groupShieldMismatch}; " +
            $"regex-ошибки группы={regexMismatchGroups}.");

        if (issues.Count > 0)
        {
            for (int i = 0; i < issues.Count; i++)
            {
                ValidationIssue issue = issues[i];
                _log.Write($"[{i + 1}] {issue.Severity} {issue.Code}: {issue.Message}");
            }

            PromptIntegerOptions pickOptions = new($"\nВыберите номер проблемы для выделения [1..{issues.Count}] или Enter: ")
            {
                AllowNone = true,
                LowerLimit = 1,
                UpperLimit = issues.Count
            };
            PromptIntegerResult pickResult = doc.Editor.GetInteger(pickOptions);
            if (pickResult.Status == PromptStatus.OK)
            {
                SelectIssueEntity(doc.Editor, issues[pickResult.Value - 1]);
            }
        }
        // END_BLOCK_COMMAND_EOM_VALIDATE
    }

    private void OnObjectAppended(ObjectId objectId)
    {
        // START_BLOCK_ON_OBJECT_APPENDED
        _xdata.ApplyActiveGroupToEntity(objectId);
        // END_BLOCK_ON_OBJECT_APPENDED
    }

    private OlsSourceSignature? PickPanelLayoutSourceSignatureFromDrawing()
    {
        // START_BLOCK_PICK_PANEL_LAYOUT_SOURCE_SIGNATURE
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return null;
        }

        var sourceOptions = new PromptEntityOptions("\nВыберите исходный блок ОЛС: ");
        sourceOptions.SetRejectMessage("\nНужен блок (BlockReference).");
        sourceOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult sourceResult = doc.Editor.GetEntity(sourceOptions);
        if (sourceResult.Status != PromptStatus.OK)
        {
            return null;
        }

        if (!TryReadSourceSignature(doc, sourceResult.ObjectId, out OlsSourceSignature signature))
        {
            _log.Write("Настройка компоновки: не удалось прочитать SOURCE (имя блока/видимость).");
            return null;
        }

        return signature;
        // END_BLOCK_PICK_PANEL_LAYOUT_SOURCE_SIGNATURE
    }

    private string? PickPanelLayoutVisualizationBlockFromDrawing()
    {
        // START_BLOCK_PICK_PANEL_LAYOUT_VISUALIZATION_BLOCK
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return null;
        }

        var targetOptions = new PromptEntityOptions("\nВыберите блок визуализации: ");
        targetOptions.SetRejectMessage("\nНужен блок (BlockReference).");
        targetOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult targetResult = doc.Editor.GetEntity(targetOptions);
        if (targetResult.Status != PromptStatus.OK)
        {
            return null;
        }

        if (!TryReadEffectiveBlockName(doc, targetResult.ObjectId, out string layoutBlockName))
        {
            _log.Write("Настройка компоновки: не удалось определить имя блока визуализации.");
            return null;
        }

        return layoutBlockName.Trim();
        // END_BLOCK_PICK_PANEL_LAYOUT_VISUALIZATION_BLOCK
    }

    private bool TryResolveExcelTemplatePath(string command, out string templatePath)
    {
        // START_BLOCK_TRY_RESOLVE_EXCEL_TEMPLATE_PATH
        templatePath = string.Empty;
        try
        {
            SettingsModel settings = _settings.LoadSettings();
            templatePath = _settings.ResolveTemplatePath(settings.ExcelTemplatePath);
            if (string.IsNullOrWhiteSpace(Path.GetFileNameWithoutExtension(templatePath)))
            {
                _log.Write($"{command}: некорректный путь шаблона Excel. Проверьте поле ExcelTemplatePath в Settings.json.");
                return false;
            }

            return true;
        }
        catch (System.Exception ex)
        {
            _log.Write($"{command}: ошибка чтения пути шаблона Excel: {ex.Message}");
            return false;
        }
        // END_BLOCK_TRY_RESOLVE_EXCEL_TEMPLATE_PATH
    }

    private static string GetExpectedExcelInputPath(string templatePath)
    {
        // START_BLOCK_GET_EXPECTED_EXCEL_INPUT_PATH
        string directory = Path.GetDirectoryName(templatePath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        return Path.Combine(string.IsNullOrWhiteSpace(directory) ? "." : directory, $"{fileName}.INPUT.csv");
        // END_BLOCK_GET_EXPECTED_EXCEL_INPUT_PATH
    }

    private static string GetExpectedExcelOutputPath(string templatePath)
    {
        // START_BLOCK_GET_EXPECTED_EXCEL_OUTPUT_PATH
        string directory = Path.GetDirectoryName(templatePath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        return Path.Combine(string.IsNullOrWhiteSpace(directory) ? "." : directory, $"{fileName}.OUTPUT.csv");
        // END_BLOCK_GET_EXPECTED_EXCEL_OUTPUT_PATH
    }

    private static ExcelInputRow ToExcelInputRow(GroupTraceAggregate row)
    {
        // START_BLOCK_TO_EXCEL_INPUT_ROW
        string group = row.Group;
        double total = row.TotalLengthMeters;
        double ceiling = GetInstallTypeLength(row, PluginConfig.Strings.Ceiling);
        double floor = GetInstallTypeLength(row, PluginConfig.Strings.Floor);
        double riser = GetInstallTypeLength(row, PluginConfig.Strings.Riser);

        return new ExcelInputRow(
            Shield: row.Shield,
            Group: group,
            PowerKw: Math.Round(row.TotalPowerWatts / 1000d, 3),
            Voltage: 0,
            TotalLengthMeters: total,
            CeilingLengthMeters: ceiling,
            FloorLengthMeters: floor,
            RiserLengthMeters: riser);
        // END_BLOCK_TO_EXCEL_INPUT_ROW
    }

    private static double GetInstallTypeLength(GroupTraceAggregate aggregate, string installType)
    {
        // START_BLOCK_GET_INSTALL_TYPE_LENGTH
        return aggregate.LengthByInstallType.TryGetValue(installType, out double value)
            ? Math.Round(value, 3)
            : 0;
        // END_BLOCK_GET_INSTALL_TYPE_LENGTH
    }

    private static IReadOnlyDictionary<string, string> ReadBlockAttributes(Transaction tr, BlockReference block)
    {
        // START_BLOCK_READ_BLOCK_ATTRIBUTES_INLINE
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (ObjectId attId in block.AttributeCollection)
        {
            if (tr.GetObject(attId, OpenMode.ForRead) is AttributeReference att)
            {
                result[att.Tag] = att.TextString ?? string.Empty;
            }
        }

        return result;
        // END_BLOCK_READ_BLOCK_ATTRIBUTES_INLINE
    }

    private static string? ReadLineGroup(Entity entity)
    {
        // START_BLOCK_READ_LINE_GROUP_INLINE
        if (entity.XData is null)
        {
            return null;
        }

        TypedValue[] values = entity.XData.AsArray();
        for (int i = 0; i < values.Length - 1; i++)
        {
            if (values[i].TypeCode != (int)DxfCode.ExtendedDataRegAppName)
            {
                continue;
            }

            string app = values[i].Value?.ToString() ?? string.Empty;
            if (!string.Equals(app, PluginConfig.Metadata.ElToolAppName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string payload = values[i + 1].Value?.ToString() ?? string.Empty;
            string[] parts = payload.Split('|');
            return parts.Length > 0 ? parts[0]?.Trim() : null;
        }

        return null;
        // END_BLOCK_READ_LINE_GROUP_INLINE
    }

    private static void DrawOlsRows(IReadOnlyList<ExcelOutputRow> rows, Point3d basePoint)
    {
        // START_BLOCK_DRAW_OLS_ROWS
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null || rows.Count == 0)
        {
            return;
        }

        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            EnsureRegApp(tr, doc.Database, PluginConfig.Metadata.OlsLayoutAppName);

            double y = basePoint.Y;
            const double step = 14.0;
            const double xInput = 0.0;
            const double xBreaker = 25.0;
            const double xRcd = 50.0;

            var title = new DBText
            {
                Position = new Point3d(basePoint.X, y, basePoint.Z),
                Height = 2.5,
                TextString = "\u041e\u0414\u041d\u041e\u041b\u0418\u041d\u0415\u0419\u041d\u0410\u042f \u0421\u0425\u0415\u041c\u0410",
                Layer = "0"
            };
            ms.AppendEntity(title);
            tr.AddNewlyCreatedDBObject(title, true);
            y -= step;

            foreach (ExcelOutputRow row in rows)
            {
                Point3d inputPoint = new(basePoint.X + xInput, y, basePoint.Z);
                Point3d breakerPoint = new(basePoint.X + xBreaker, y, basePoint.Z);
                Point3d rcdPoint = new(basePoint.X + xRcd, y, basePoint.Z);

                InsertTemplateOrFallback(tr, bt, ms, PluginConfig.TemplateBlocks.Input, inputPoint, "\u0412\u0412\u041e\u0414");
                InsertTemplateOrFallback(tr, bt, ms, PluginConfig.TemplateBlocks.Breaker, breakerPoint, row.CircuitBreaker);
                InsertTemplateOrFallback(tr, bt, ms, PluginConfig.TemplateBlocks.Rcd, rcdPoint, row.RcdDiff);

                var l1 = new Line(new Point3d(inputPoint.X + 5, inputPoint.Y, inputPoint.Z), new Point3d(breakerPoint.X - 1, breakerPoint.Y, breakerPoint.Z));
                ms.AppendEntity(l1);
                tr.AddNewlyCreatedDBObject(l1, true);

                var l2 = new Line(new Point3d(breakerPoint.X + 5, breakerPoint.Y, breakerPoint.Z), new Point3d(rcdPoint.X - 1, rcdPoint.Y, rcdPoint.Z));
                ms.AppendEntity(l2);
                tr.AddNewlyCreatedDBObject(l2, true);

                var label = new DBText
                {
                    Position = new Point3d(basePoint.X + xRcd + 12, y, basePoint.Z),
                    Height = 2.5,
                    TextString = $"{row.Group} | {row.Cable} | {row.CircuitBreaker}",
                    Layer = "0"
                };
                ms.AppendEntity(label);
                tr.AddNewlyCreatedDBObject(label, true);
                WriteOlsRowMetadata(label, row);

                y -= step;
            }

            tr.Commit();
        }
        // END_BLOCK_DRAW_OLS_ROWS
    }

    private static void WriteOlsRowMetadata(Entity entity, ExcelOutputRow row)
    {
        // START_BLOCK_WRITE_OLS_ROW_METADATA
        UpsertEntityXData(entity, PluginConfig.Metadata.OlsLayoutAppName, [
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, "OLS_ROW"),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, row.Shield ?? string.Empty),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, row.Group ?? string.Empty),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, row.CircuitBreaker ?? string.Empty),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, row.RcdDiff ?? string.Empty),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, row.Cable ?? string.Empty),
            new TypedValue((int)DxfCode.ExtendedDataInteger16, (short)Math.Clamp(row.CircuitBreakerModules, 0, short.MaxValue)),
            new TypedValue((int)DxfCode.ExtendedDataInteger16, (short)Math.Clamp(row.RcdModules, 0, short.MaxValue)),
            new TypedValue((int)DxfCode.ExtendedDataAsciiString, row.Note ?? string.Empty)
        ]);
        // END_BLOCK_WRITE_OLS_ROW_METADATA
    }

    private static IReadOnlyList<ExcelOutputRow> ReadOlsRowsFromDrawing(Document doc)
    {
        // START_BLOCK_READ_OLS_ROWS_FROM_DRAWING
        var rows = new List<ExcelOutputRow>();
        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            BlockTable blockTable = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            foreach (ObjectId id in modelSpace)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is not Entity entity)
                {
                    continue;
                }

                ExcelOutputRow? row = TryParseOlsRowFromEntity(entity);
                if (row is not null)
                {
                    rows.Add(row);
                }
            }

            tr.Commit();
        }

        return rows;
        // END_BLOCK_READ_OLS_ROWS_FROM_DRAWING
    }

    private static ExcelOutputRow? TryParseOlsRowFromEntity(Entity entity)
    {
        // START_BLOCK_TRY_PARSE_OLS_ROW_FROM_ENTITY
        if (entity.XData is null)
        {
            return null;
        }

        TypedValue[] values = entity.XData.AsArray();
        for (int i = 0; i < values.Length; i++)
        {
            if (values[i].TypeCode != (int)DxfCode.ExtendedDataRegAppName)
            {
                continue;
            }

            string appName = values[i].Value?.ToString() ?? string.Empty;
            if (!string.Equals(appName, PluginConfig.Metadata.OlsLayoutAppName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (i + 8 >= values.Length)
            {
                return null;
            }

            string marker = values[i + 1].Value?.ToString() ?? string.Empty;
            if (!string.Equals(marker, "OLS_ROW", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return new ExcelOutputRow(
                Shield: values[i + 2].Value?.ToString() ?? string.Empty,
                Group: values[i + 3].Value?.ToString() ?? string.Empty,
                CircuitBreaker: values[i + 4].Value?.ToString() ?? string.Empty,
                RcdDiff: values[i + 5].Value?.ToString() ?? string.Empty,
                Cable: values[i + 6].Value?.ToString() ?? string.Empty,
                CircuitBreakerModules: ToIntOrDefault(values[i + 7].Value),
                RcdModules: ToIntOrDefault(values[i + 8].Value),
                Note: i + 9 < values.Length ? values[i + 9].Value?.ToString() : null);
        }

        return null;
        // END_BLOCK_TRY_PARSE_OLS_ROW_FROM_ENTITY
    }

    private static int ToIntOrDefault(object? value)
    {
        // START_BLOCK_TO_INT_OR_DEFAULT
        if (value is short shortValue)
        {
            return shortValue;
        }

        if (value is int intValue)
        {
            return intValue;
        }

        return int.TryParse(value?.ToString(), out int parsed) ? parsed : 0;
        // END_BLOCK_TO_INT_OR_DEFAULT
    }

    private static ParsedOlsSelection ParseOlsSelection(
        Document doc,
        IReadOnlyList<ObjectId> selectedIds,
        PanelLayoutAttributeTags tags)
    {
        // START_BLOCK_PARSE_OLS_SELECTION
        var devices = new List<OlsSelectedDevice>();
        var issues = new List<SkippedOlsDeviceIssue>();

        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            foreach (ObjectId id in selectedIds)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is not BlockReference block)
                {
                    issues.Add(new SkippedOlsDeviceIssue(id, "Объект не является BlockReference."));
                    continue;
                }

                IReadOnlyDictionary<string, string> attrs = ReadBlockAttributes(tr, block);
                attrs.TryGetValue(tags.Device, out string? rawDevice);
                attrs.TryGetValue(tags.Modules, out string? rawModules);
                int modules = TryParsePositiveInt(rawModules, out int parsedModules) ? parsedModules : 0;

                string sourceBlockName = ResolveEffectiveBlockName(tr, block);
                if (string.IsNullOrWhiteSpace(sourceBlockName))
                {
                    issues.Add(new SkippedOlsDeviceIssue(id, "Не удалось определить имя исходного блока ОЛС."));
                    continue;
                }

                string? visibilityValue = ResolveVisibilityValue(block);
                attrs.TryGetValue(tags.Group, out string? group);
                attrs.TryGetValue(tags.Note, out string? note);
                devices.Add(new OlsSelectedDevice(
                    EntityId: id,
                    SourceBlockName: sourceBlockName,
                    SourceSignature: new OlsSourceSignature(sourceBlockName, visibilityValue),
                    DeviceKey: string.IsNullOrWhiteSpace(rawDevice) ? null : rawDevice.Trim(),
                    Modules: modules,
                    Group: string.IsNullOrWhiteSpace(group) ? null : group.Trim(),
                    Note: string.IsNullOrWhiteSpace(note) ? null : note.Trim(),
                    InsertionPoint: block.Position));
            }

            tr.Commit();
        }

        IReadOnlyList<OlsSelectedDevice> sorted = devices
            .OrderByDescending(x => x.InsertionPoint.Y)
            .ThenBy(x => x.InsertionPoint.X)
            .ToList();
        return new ParsedOlsSelection(sorted, issues);
        // END_BLOCK_PARSE_OLS_SELECTION
    }

    private IReadOnlyList<MappedLayoutDevice> MapOlsDevices(
        IReadOnlyList<OlsSelectedDevice> devices,
        IReadOnlyList<PanelLayoutSelectorRule> selectorRules,
        IReadOnlyList<PanelLayoutMapRule> legacyRules,
        out List<SkippedOlsDeviceIssue> issues)
    {
        // START_BLOCK_MAP_OLS_DEVICES
        issues = new List<SkippedOlsDeviceIssue>();
        var result = new List<MappedLayoutDevice>(devices.Count);
        IReadOnlyList<PanelLayoutSelectorRule> orderedSelectorRules = selectorRules
            .Where(x => x is not null)
            .Where(x => !string.IsNullOrWhiteSpace(x.SourceBlockName) && !string.IsNullOrWhiteSpace(x.LayoutBlockName))
            .Select(x => new PanelLayoutSelectorRule(
                x.Priority < 0 ? 0 : x.Priority,
                x.SourceBlockName.Trim(),
                string.IsNullOrWhiteSpace(x.VisibilityValue) ? null : x.VisibilityValue.Trim(),
                x.LayoutBlockName.Trim(),
                x.FallbackModules is > 0 ? x.FallbackModules : null))
            .OrderBy(x => x.Priority)
            .ThenBy(x => x.SourceBlockName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.VisibilityValue ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var legacyMap = legacyRules
            .Where(x => !string.IsNullOrWhiteSpace(x.DeviceKey) && !string.IsNullOrWhiteSpace(x.LayoutBlockName))
            .GroupBy(x => x.DeviceKey.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);

        foreach (OlsSelectedDevice device in devices)
        {
            string displayLabel = ResolveDeviceLabel(device);
            PanelLayoutSelectorRule? selectorRule = TryResolveSelectorRule(device.SourceSignature, orderedSelectorRules, out int samePriorityMatches);

            string? layoutBlockName = null;
            int? fallbackModules = null;
            if (selectorRule is not null)
            {
                layoutBlockName = selectorRule.LayoutBlockName;
                fallbackModules = selectorRule.FallbackModules;
                if (samePriorityMatches > 1)
                {
                    _log.Write($"EOM_КОМПОНОВКА_ЩИТА: SOURCE='{FormatSourceSignature(device.SourceSignature)}' совпадает с несколькими правилами Priority={selectorRule.Priority}. Применено первое правило.");
                }
            }
            else if (!string.IsNullOrWhiteSpace(device.DeviceKey) && legacyMap.TryGetValue(device.DeviceKey.Trim(), out PanelLayoutMapRule? legacyRule))
            {
                layoutBlockName = legacyRule.LayoutBlockName.Trim();
                fallbackModules = legacyRule.FallbackModules;
            }

            if (selectorRule is not null &&
                !fallbackModules.HasValue &&
                !string.IsNullOrWhiteSpace(device.DeviceKey) &&
                legacyMap.TryGetValue(device.DeviceKey.Trim(), out PanelLayoutMapRule? legacyForModules))
            {
                fallbackModules = legacyForModules.FallbackModules;
            }

            if (string.IsNullOrWhiteSpace(layoutBlockName))
            {
                if (string.IsNullOrWhiteSpace(device.DeviceKey) && device.Modules <= 0)
                {
                    continue;
                }

                issues.Add(new SkippedOlsDeviceIssue(
                    device.EntityId,
                    $"Нет правила PanelLayoutMap.json для SOURCE='{FormatSourceSignature(device.SourceSignature)}' (и fallback по АППАРАТ не найден).",
                    device.DeviceKey,
                    device.SourceBlockName));
                continue;
            }

            int resolvedModules = ResolveModuleCount(device.Modules, fallbackModules);
            if (resolvedModules <= 0)
            {
                issues.Add(new SkippedOlsDeviceIssue(
                    device.EntityId,
                    $"Отсутствует корректный атрибут {PluginConfig.PanelLayout.ModulesTag} и в правиле нет FallbackModules.",
                    device.DeviceKey,
                    device.SourceBlockName));
                continue;
            }

            result.Add(new MappedLayoutDevice(
                EntityId: device.EntityId,
                SourceBlockName: device.SourceBlockName,
                DeviceKey: device.DeviceKey,
                DisplayLabel: displayLabel,
                LayoutBlockName: layoutBlockName.Trim(),
                Modules: resolvedModules,
                Group: device.Group,
                Note: device.Note));
        }

        return result;
        // END_BLOCK_MAP_OLS_DEVICES
    }

    private static int ResolveModuleCount(int modulesFromAttribute, int? fallbackModules)
    {
        // START_BLOCK_RESOLVE_MODULE_COUNT
        if (modulesFromAttribute > 0)
        {
            return modulesFromAttribute;
        }

        return fallbackModules is > 0 ? fallbackModules.Value : 0;
        // END_BLOCK_RESOLVE_MODULE_COUNT
    }

    private static string ResolveDeviceLabel(OlsSelectedDevice device)
    {
        // START_BLOCK_RESOLVE_DEVICE_LABEL
        if (!string.IsNullOrWhiteSpace(device.DeviceKey))
        {
            return device.DeviceKey.Trim();
        }

        return FormatSourceSignature(device.SourceSignature);
        // END_BLOCK_RESOLVE_DEVICE_LABEL
    }

    private static PanelLayoutSelectorRule? TryResolveSelectorRule(
        OlsSourceSignature signature,
        IReadOnlyList<PanelLayoutSelectorRule> rules,
        out int samePriorityMatches)
    {
        // START_BLOCK_TRY_RESOLVE_SELECTOR_RULE
        samePriorityMatches = 0;
        PanelLayoutSelectorRule? selected = null;
        int bestPriority = int.MaxValue;
        int bestSpecificity = -1;

        foreach (PanelLayoutSelectorRule rule in rules)
        {
            if (!IsSelectorRuleMatch(rule, signature))
            {
                continue;
            }

            int specificity = GetSelectorRuleSpecificity(rule);
            if (selected is null)
            {
                selected = rule;
                bestPriority = rule.Priority;
                bestSpecificity = specificity;
                samePriorityMatches = 1;
                continue;
            }

            if (rule.Priority < bestPriority ||
                (rule.Priority == bestPriority && specificity > bestSpecificity))
            {
                selected = rule;
                bestPriority = rule.Priority;
                bestSpecificity = specificity;
                samePriorityMatches = 1;
                continue;
            }

            if (rule.Priority == bestPriority && specificity == bestSpecificity)
            {
                samePriorityMatches++;
            }
        }

        return selected;
        // END_BLOCK_TRY_RESOLVE_SELECTOR_RULE
    }

    private static int GetSelectorRuleSpecificity(PanelLayoutSelectorRule rule)
    {
        // START_BLOCK_GET_SELECTOR_RULE_SPECIFICITY
        return string.IsNullOrWhiteSpace(rule.VisibilityValue) ? 0 : 1;
        // END_BLOCK_GET_SELECTOR_RULE_SPECIFICITY
    }

    private static bool IsSelectorRuleMatch(PanelLayoutSelectorRule rule, OlsSourceSignature signature)
    {
        // START_BLOCK_IS_SELECTOR_RULE_MATCH
        if (!string.Equals(rule.SourceBlockName, signature.SourceBlockName, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(rule.VisibilityValue))
        {
            return true;
        }

        return string.Equals(rule.VisibilityValue, signature.VisibilityValue, StringComparison.OrdinalIgnoreCase);
        // END_BLOCK_IS_SELECTOR_RULE_MATCH
    }

    private static string FormatSourceSignature(OlsSourceSignature signature)
    {
        // START_BLOCK_FORMAT_SOURCE_SIGNATURE
        if (string.IsNullOrWhiteSpace(signature.VisibilityValue))
        {
            return $"{signature.SourceBlockName}|*";
        }

        return $"{signature.SourceBlockName}|{signature.VisibilityValue}";
        // END_BLOCK_FORMAT_SOURCE_SIGNATURE
    }

    // START_CONTRACT: BuildPanelLayoutModel
    //   PURPOSE: Build split-aware DIN row placement model for EOM_КОМПОНОВКА_ЩИТА from mapped OLS devices.
    //   INPUTS: { devices: IReadOnlyList<MappedLayoutDevice> - mapped OLS devices, modulesPerRow: int - requested modules in one DIN row }
    //   OUTPUTS: { IReadOnlyList<PanelLayoutRow> - normalized placements with row/slot coordinates and split metadata }
    // END_CONTRACT: BuildPanelLayoutModel
    private static IReadOnlyList<PanelLayoutRow> BuildPanelLayoutModel(IReadOnlyList<MappedLayoutDevice> devices, int modulesPerRow)
    {
        // START_BLOCK_BUILD_PANEL_LAYOUT_MODEL
        if (devices.Count == 0)
        {
            return [];
        }

        int safeModulesPerRow = modulesPerRow <= 0 ? 24 : modulesPerRow;
        var result = new List<PanelLayoutRow>(devices.Count);
        int dinRow = 1;
        int occupiedInRow = 0;

        foreach (MappedLayoutDevice device in devices)
        {
            int totalModules = NormalizeModuleCount(device.Modules);
            int remainingModules = totalModules;
            int segmentCount = CountSegmentsForDevice(totalModules, occupiedInRow, safeModulesPerRow);
            int segmentIndex = 1;

            while (remainingModules > 0)
            {
                if (occupiedInRow >= safeModulesPerRow)
                {
                    dinRow++;
                    occupiedInRow = 0;
                }

                int freeSlots = safeModulesPerRow - occupiedInRow;
                int modulesInSegment = Math.Min(remainingModules, freeSlots);
                int slotStart = occupiedInRow + 1;
                int slotEnd = slotStart + modulesInSegment - 1;

                result.Add(new PanelLayoutRow(
                    EntityId: device.EntityId,
                    DinRow: dinRow,
                    SlotStart: slotStart,
                    SlotEnd: slotEnd,
                    LayoutBlockName: device.LayoutBlockName,
                    DeviceKey: device.DisplayLabel,
                    SourceBlockName: device.SourceBlockName,
                    ModuleCount: modulesInSegment,
                    SegmentIndex: segmentIndex,
                    SegmentCount: segmentCount,
                    Group: device.Group,
                    Note: device.Note));

                occupiedInRow += modulesInSegment;
                remainingModules -= modulesInSegment;
                segmentIndex++;
            }
        }

        return result;
        // END_BLOCK_BUILD_PANEL_LAYOUT_MODEL
    }

    private static int NormalizeModuleCount(int rawModuleCount)
    {
        // START_BLOCK_NORMALIZE_MODULE_COUNT
        return rawModuleCount <= 0 ? 1 : rawModuleCount;
        // END_BLOCK_NORMALIZE_MODULE_COUNT
    }

    private static int CountSegmentsForDevice(int totalModules, int occupiedInRow, int modulesPerRow)
    {
        // START_BLOCK_COUNT_SEGMENTS_FOR_DEVICE
        int remaining = totalModules;
        int occupied = occupiedInRow;
        int segments = 0;
        while (remaining > 0)
        {
            if (occupied >= modulesPerRow)
            {
                occupied = 0;
            }

            int freeSlots = modulesPerRow - occupied;
            int chunk = Math.Min(remaining, freeSlots);
            remaining -= chunk;
            occupied += chunk;
            segments++;
        }

        return Math.Max(1, segments);
        // END_BLOCK_COUNT_SEGMENTS_FOR_DEVICE
    }

    private static int DrawPanelLayout(
        IReadOnlyList<PanelLayoutRow> layoutRows,
        Point3d basePoint,
        int modulesPerRow,
        out List<SkippedOlsDeviceIssue> issues)
    {
        // START_BLOCK_DRAW_PANEL_LAYOUT_MODEL
        issues = new List<SkippedOlsDeviceIssue>();
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null || layoutRows.Count == 0)
        {
            return 0;
        }

        int safeModulesPerRow = modulesPerRow <= 0 ? 24 : modulesPerRow;
        const double moduleWidth = 8.0;
        const double moduleHeight = 6.0;
        const double rowGap = 3.0;
        int renderedSegments = 0;

        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            int maxDinRow = layoutRows.Max(x => x.DinRow);
            for (int dinRow = 1; dinRow <= maxDinRow; dinRow++)
            {
                double rowTop = basePoint.Y - ((dinRow - 1) * (moduleHeight + rowGap));
                DrawPanelLayoutRowFrame(ms, tr, basePoint.X, rowTop, basePoint.Z, safeModulesPerRow, moduleWidth, moduleHeight, dinRow);
            }

            foreach (PanelLayoutRow row in layoutRows.OrderBy(x => x.DinRow).ThenBy(x => x.SlotStart))
            {
                double rowTop = basePoint.Y - ((row.DinRow - 1) * (moduleHeight + rowGap));
                double leftX = basePoint.X + ((row.SlotStart - 1) * moduleWidth);
                bool drawn = DrawPanelLayoutDevice(ms, tr, bt, row, leftX, rowTop, basePoint.Z, moduleWidth, moduleHeight);
                if (!drawn)
                {
                    issues.Add(new SkippedOlsDeviceIssue(
                        row.EntityId,
                        $"В чертеже отсутствует блок визуализации '{row.LayoutBlockName}'.",
                        row.DeviceKey,
                        row.SourceBlockName));
                    continue;
                }

                renderedSegments++;
            }

            tr.Commit();
        }

        return renderedSegments;
        // END_BLOCK_DRAW_PANEL_LAYOUT_MODEL
    }

    private static void DrawPanelLayoutRowFrame(
        BlockTableRecord ms,
        Transaction tr,
        double rowStartX,
        double rowTopY,
        double z,
        int modulesPerRow,
        double moduleWidth,
        double moduleHeight,
        int dinRow)
    {
        // START_BLOCK_DRAW_PANEL_LAYOUT_ROW_FRAME
        DrawRectangle(ms, tr, rowStartX, rowTopY, moduleWidth * modulesPerRow, moduleHeight);

        var rowLabel = new DBText
        {
            Position = new Point3d(rowStartX - 12.0, rowTopY - (moduleHeight / 2.0), z),
            Height = 2.0,
            TextString = $"Ряд {dinRow}",
            Layer = "0"
        };
        ms.AppendEntity(rowLabel);
        tr.AddNewlyCreatedDBObject(rowLabel, true);
        // END_BLOCK_DRAW_PANEL_LAYOUT_ROW_FRAME
    }

    private static bool DrawPanelLayoutDevice(
        BlockTableRecord ms,
        Transaction tr,
        BlockTable bt,
        PanelLayoutRow row,
        double leftX,
        double topY,
        double z,
        double moduleWidth,
        double moduleHeight)
    {
        // START_BLOCK_DRAW_PANEL_LAYOUT_DEVICE
        if (!bt.Has(row.LayoutBlockName))
        {
            return false;
        }

        double width = moduleWidth * row.ModuleCount;
        DrawRectangle(ms, tr, leftX, topY, width, moduleHeight);
        if (!TryInsertLayoutBlockFitted(ms, tr, bt, row.LayoutBlockName, leftX, topY, z, width, moduleHeight))
        {
            return false;
        }

        string segmentSuffix = row.SegmentCount > 1
            ? $" ({row.SegmentIndex}/{row.SegmentCount})"
            : string.Empty;
        string groupPrefix = string.IsNullOrWhiteSpace(row.Group) ? string.Empty : $"{row.Group}: ";
        var label = new DBText
        {
            Position = new Point3d(leftX + 0.5, topY - moduleHeight - 1.8, z),
            Height = 1.8,
            TextString = $"{groupPrefix}{row.DeviceKey}{segmentSuffix} [{row.ModuleCount}]",
            Layer = "0"
        };
        ms.AppendEntity(label);
        tr.AddNewlyCreatedDBObject(label, true);
        return true;
        // END_BLOCK_DRAW_PANEL_LAYOUT_DEVICE
    }

    private static bool TryInsertLayoutBlockFitted(
        BlockTableRecord ms,
        Transaction tr,
        BlockTable bt,
        string blockName,
        double leftX,
        double topY,
        double z,
        double targetWidth,
        double targetHeight)
    {
        // START_BLOCK_TRY_INSERT_LAYOUT_BLOCK_FITTED
        if (!bt.Has(blockName))
        {
            return false;
        }

        ObjectId blockId = bt[blockName];
        if (tr.GetObject(blockId, OpenMode.ForRead) is not BlockTableRecord btr)
        {
            return false;
        }

        if (!TryGetBlockDefinitionExtents(tr, btr, out Point3d sourceMin, out Point3d sourceMax))
        {
            var fallbackRef = new BlockReference(
                new Point3d(leftX + (targetWidth / 2.0), topY - (targetHeight / 2.0), z),
                blockId);
            ms.AppendEntity(fallbackRef);
            tr.AddNewlyCreatedDBObject(fallbackRef, true);
            return true;
        }

        double sourceWidth = Math.Max(0.001, sourceMax.X - sourceMin.X);
        double scale = 1.0;
        // Keep legacy visual size (scale=1) and only shrink if block is wider than slot.
        if (sourceWidth > targetWidth && targetWidth > 0.0)
        {
            scale = targetWidth / sourceWidth;
        }

        Point3d sourceCenter = new(
            (sourceMin.X + sourceMax.X) / 2.0,
            (sourceMin.Y + sourceMax.Y) / 2.0,
            (sourceMin.Z + sourceMax.Z) / 2.0);
        Point3d targetCenter = new(leftX + (targetWidth / 2.0), topY - (targetHeight / 2.0), z);
        Point3d insertionPoint = new(
            targetCenter.X - (sourceCenter.X * scale),
            targetCenter.Y - (sourceCenter.Y * scale),
            targetCenter.Z - (sourceCenter.Z * scale));

        var blockRef = new BlockReference(insertionPoint, blockId)
        {
            ScaleFactors = new Scale3d(scale)
        };
        ms.AppendEntity(blockRef);
        tr.AddNewlyCreatedDBObject(blockRef, true);
        return true;
        // END_BLOCK_TRY_INSERT_LAYOUT_BLOCK_FITTED
    }

    private static bool TryGetBlockDefinitionExtents(
        Transaction tr,
        BlockTableRecord btr,
        out Point3d minPoint,
        out Point3d maxPoint)
    {
        // START_BLOCK_TRY_GET_BLOCK_DEFINITION_EXTENTS
        minPoint = Point3d.Origin;
        maxPoint = Point3d.Origin;

        bool hasExtents = false;
        double minX = double.MaxValue;
        double minY = double.MaxValue;
        double minZ = double.MaxValue;
        double maxX = double.MinValue;
        double maxY = double.MinValue;
        double maxZ = double.MinValue;

        foreach (ObjectId id in btr)
        {
            if (tr.GetObject(id, OpenMode.ForRead) is not Entity entity)
            {
                continue;
            }

            Extents3d extents;
            try
            {
                extents = entity.GeometricExtents;
            }
            catch
            {
                continue;
            }

            hasExtents = true;
            minX = Math.Min(minX, extents.MinPoint.X);
            minY = Math.Min(minY, extents.MinPoint.Y);
            minZ = Math.Min(minZ, extents.MinPoint.Z);
            maxX = Math.Max(maxX, extents.MaxPoint.X);
            maxY = Math.Max(maxY, extents.MaxPoint.Y);
            maxZ = Math.Max(maxZ, extents.MaxPoint.Z);
        }

        if (!hasExtents)
        {
            return false;
        }

        minPoint = new Point3d(minX, minY, minZ);
        maxPoint = new Point3d(maxX, maxY, maxZ);
        return true;
        // END_BLOCK_TRY_GET_BLOCK_DEFINITION_EXTENTS
    }

    private static void DrawRectangle(BlockTableRecord ms, Transaction tr, double leftX, double topY, double width, double height)
    {
        // START_BLOCK_DRAW_RECTANGLE_ENTITY
        var frame = new Polyline();
        frame.AddVertexAt(0, new Point2d(leftX, topY), 0, 0, 0);
        frame.AddVertexAt(1, new Point2d(leftX + width, topY), 0, 0, 0);
        frame.AddVertexAt(2, new Point2d(leftX + width, topY - height), 0, 0, 0);
        frame.AddVertexAt(3, new Point2d(leftX, topY - height), 0, 0, 0);
        frame.Closed = true;
        ms.AppendEntity(frame);
        tr.AddNewlyCreatedDBObject(frame, true);
        // END_BLOCK_DRAW_RECTANGLE_ENTITY
    }

    private static string ResolveBlockName(Transaction tr, BlockReference block)
    {
        // START_BLOCK_RESOLVE_BLOCK_NAME
        if (tr.GetObject(block.BlockTableRecord, OpenMode.ForRead) is BlockTableRecord btr)
        {
            return btr.Name ?? string.Empty;
        }

        return string.Empty;
        // END_BLOCK_RESOLVE_BLOCK_NAME
    }

    private static string ResolveEffectiveBlockName(Transaction tr, BlockReference block)
    {
        // START_BLOCK_RESOLVE_EFFECTIVE_BLOCK_NAME
        if (block.IsDynamicBlock && !block.DynamicBlockTableRecord.IsNull)
        {
            if (tr.GetObject(block.DynamicBlockTableRecord, OpenMode.ForRead) is BlockTableRecord dynamicBtr &&
                !string.IsNullOrWhiteSpace(dynamicBtr.Name) &&
                !IsAnonymousBlockName(dynamicBtr.Name))
            {
                return dynamicBtr.Name.Trim();
            }
        }

        string rawName = ResolveBlockName(tr, block);
        if (!string.IsNullOrWhiteSpace(rawName) && !IsAnonymousBlockName(rawName))
        {
            return rawName.Trim();
        }

        string? ownerName = TryResolveOwnerDynamicBlockName(tr, block);
        if (!string.IsNullOrWhiteSpace(ownerName))
        {
            return ownerName.Trim();
        }

        return rawName;
        // END_BLOCK_RESOLVE_EFFECTIVE_BLOCK_NAME
    }

    private static string? TryResolveOwnerDynamicBlockName(Transaction tr, BlockReference block)
    {
        // START_BLOCK_TRY_RESOLVE_OWNER_DYNAMIC_BLOCK_NAME
        Database? db = block.Database;
        if (db is null || block.BlockTableRecord.IsNull)
        {
            return null;
        }

        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        foreach (ObjectId candidateId in bt)
        {
            if (tr.GetObject(candidateId, OpenMode.ForRead) is not BlockTableRecord candidate)
            {
                continue;
            }

            if (candidate.IsAnonymous || string.IsNullOrWhiteSpace(candidate.Name))
            {
                continue;
            }

            ObjectIdCollection anonymousIds;
            try
            {
                anonymousIds = candidate.GetAnonymousBlockIds();
            }
            catch
            {
                continue;
            }

            foreach (ObjectId anonId in anonymousIds)
            {
                if (anonId == block.BlockTableRecord)
                {
                    return candidate.Name.Trim();
                }
            }
        }

        return null;
        // END_BLOCK_TRY_RESOLVE_OWNER_DYNAMIC_BLOCK_NAME
    }

    private static bool IsAnonymousBlockName(string? name)
    {
        // START_BLOCK_IS_ANONYMOUS_BLOCK_NAME
        return !string.IsNullOrWhiteSpace(name) && name.StartsWith("*", StringComparison.Ordinal);
        // END_BLOCK_IS_ANONYMOUS_BLOCK_NAME
    }

    private static string? ResolveVisibilityValue(BlockReference block)
    {
        // START_BLOCK_RESOLVE_VISIBILITY_VALUE
        if (!block.IsDynamicBlock)
        {
            return null;
        }

        try
        {
            DynamicBlockReferencePropertyCollection properties = block.DynamicBlockReferencePropertyCollection;
            foreach (DynamicBlockReferenceProperty property in properties)
            {
                if (!IsVisibilityLikeProperty(property.PropertyName))
                {
                    continue;
                }

                string? value = property.Value?.ToString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }
        }
        catch
        {
            return null;
        }

        return null;
        // END_BLOCK_RESOLVE_VISIBILITY_VALUE
    }

    private static bool IsVisibilityLikeProperty(string? propertyName)
    {
        // START_BLOCK_IS_VISIBILITY_LIKE_PROPERTY
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return false;
        }

        return propertyName.Contains("visibility", StringComparison.OrdinalIgnoreCase)
            || propertyName.Contains("видим", StringComparison.OrdinalIgnoreCase)
            || propertyName.Contains("lookup", StringComparison.OrdinalIgnoreCase)
            || propertyName.Contains("таблиц", StringComparison.OrdinalIgnoreCase);
        // END_BLOCK_IS_VISIBILITY_LIKE_PROPERTY
    }

    private static bool TryReadSourceSignature(Document doc, ObjectId objectId, out OlsSourceSignature signature)
    {
        // START_BLOCK_TRY_READ_SOURCE_SIGNATURE
        signature = new OlsSourceSignature(string.Empty, null);
        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            if (tr.GetObject(objectId, OpenMode.ForRead) is not BlockReference block)
            {
                return false;
            }

            string sourceBlockName = ResolveEffectiveBlockName(tr, block);
            if (string.IsNullOrWhiteSpace(sourceBlockName))
            {
                return false;
            }

            signature = new OlsSourceSignature(sourceBlockName, ResolveVisibilityValue(block));
            tr.Commit();
            return true;
        }
        // END_BLOCK_TRY_READ_SOURCE_SIGNATURE
    }

    private static bool TryReadEffectiveBlockName(Document doc, ObjectId objectId, out string blockName)
    {
        // START_BLOCK_TRY_READ_EFFECTIVE_BLOCK_NAME
        blockName = string.Empty;
        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            if (tr.GetObject(objectId, OpenMode.ForRead) is not BlockReference block)
            {
                return false;
            }

            blockName = ResolveEffectiveBlockName(tr, block);
            if (string.IsNullOrWhiteSpace(blockName))
            {
                return false;
            }

            tr.Commit();
            return true;
        }
        // END_BLOCK_TRY_READ_EFFECTIVE_BLOCK_NAME
    }

    private static bool TryParsePositiveInt(string? raw, out int value)
    {
        // START_BLOCK_TRY_PARSE_POSITIVE_INT
        value = 0;
        if (string.IsNullOrWhiteSpace(raw))
        {
            return false;
        }

        if (!int.TryParse(raw.Trim(), out int parsed))
        {
            return false;
        }

        if (parsed <= 0)
        {
            return false;
        }

        value = parsed;
        return true;
        // END_BLOCK_TRY_PARSE_POSITIVE_INT
    }

    private void ReportSkippedOlsDevices(IReadOnlyList<SkippedOlsDeviceIssue> issues)
    {
        // START_BLOCK_REPORT_SKIPPED_OLS_DEVICES
        if (issues.Count == 0)
        {
            return;
        }

        _log.Write($"EOM_КОМПОНОВКА_ЩИТА: пропущено объектов: {issues.Count}.");
        foreach (SkippedOlsDeviceIssue issue in issues)
        {
            string blockNamePart = string.IsNullOrWhiteSpace(issue.SourceBlockName) ? string.Empty : $" BLOCK={issue.SourceBlockName};";
            string deviceKeyPart = string.IsNullOrWhiteSpace(issue.DeviceKey) ? string.Empty : $" АППАРАТ={issue.DeviceKey};";
            _log.Write($"  - ID={issue.EntityId};{blockNamePart}{deviceKeyPart} Причина: {issue.Reason}");
        }
        // END_BLOCK_REPORT_SKIPPED_OLS_DEVICES
    }

    private sealed record ParsedOlsSelection(IReadOnlyList<OlsSelectedDevice> Devices, IReadOnlyList<SkippedOlsDeviceIssue> Issues);
    private sealed record MappedLayoutDevice(
        ObjectId EntityId,
        string SourceBlockName,
        string? DeviceKey,
        string DisplayLabel,
        string LayoutBlockName,
        int Modules,
        string? Group,
        string? Note);
    private sealed record PanelLayoutRow(
        ObjectId EntityId,
        int DinRow,
        int SlotStart,
        int SlotEnd,
        string LayoutBlockName,
        string DeviceKey,
        string SourceBlockName,
        int ModuleCount,
        int SegmentIndex,
        int SegmentCount,
        string? Group,
        string? Note);

    private static void InsertTemplateOrFallback(Transaction tr, BlockTable bt, BlockTableRecord ms, string blockName, Point3d position, string fallbackLabel)
    {
        // START_BLOCK_INSERT_TEMPLATE_OR_FALLBACK
        if (bt.Has(blockName))
        {
            ObjectId blockId = bt[blockName];
            var blockRef = new BlockReference(position, blockId);
            ms.AppendEntity(blockRef);
            tr.AddNewlyCreatedDBObject(blockRef, true);
            return;
        }

        var fallback = new DBText
        {
            Position = position,
            Height = 2.5,
            TextString = fallbackLabel,
            Layer = "0"
        };
        ms.AppendEntity(fallback);
        tr.AddNewlyCreatedDBObject(fallback, true);
        // END_BLOCK_INSERT_TEMPLATE_OR_FALLBACK
    }

    private static void EnsureRegApp(Transaction tr, Database db, string appName)
    {
        // START_BLOCK_ENSURE_REGAPP_INLINE
        RegAppTable regAppTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
        if (regAppTable.Has(appName))
        {
            return;
        }

        regAppTable.UpgradeOpen();
        var record = new RegAppTableRecord { Name = appName };
        regAppTable.Add(record);
        tr.AddNewlyCreatedDBObject(record, true);
        // END_BLOCK_ENSURE_REGAPP_INLINE
    }

    private static void UpsertEntityXData(Entity entity, string appName, IReadOnlyList<TypedValue> appPayload)
    {
        // START_BLOCK_UPSERT_ENTITY_XDATA
        List<TypedValue> allValues = entity.XData?.AsArray().ToList() ?? [];
        var merged = new List<TypedValue>();
        bool skipCurrentAppPayload = false;
        foreach (TypedValue value in allValues)
        {
            if (value.TypeCode == (int)DxfCode.ExtendedDataRegAppName)
            {
                string currentApp = value.Value?.ToString() ?? string.Empty;
                skipCurrentAppPayload = string.Equals(currentApp, appName, StringComparison.OrdinalIgnoreCase);
                if (!skipCurrentAppPayload)
                {
                    merged.Add(value);
                }

                continue;
            }

            if (!skipCurrentAppPayload)
            {
                merged.Add(value);
            }
        }

        merged.Add(new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName));
        merged.AddRange(appPayload);
        entity.XData = new ResultBuffer(merged.ToArray());
        // END_BLOCK_UPSERT_ENTITY_XDATA
    }

    private static void SelectIssueEntity(Editor editor, ValidationIssue issue)
    {
        // START_BLOCK_SELECT_ISSUE_ENTITY
        if (issue.EntityId.IsNull)
        {
            return;
        }

        editor.SetImpliedSelection(new[] { issue.EntityId });
        editor.WriteMessage($"\nВыбран объект проблемы: {issue.Code}");
        // END_BLOCK_SELECT_ISSUE_ENTITY
    }
}
