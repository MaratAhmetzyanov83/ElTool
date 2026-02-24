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

        IReadOnlyList<SpecificationRow> rows = _spec.BuildSpecification(Array.Empty<ObjectId>());
        _export.ToAutoCadTable(rows);
        _export.ToCsv(rows);
        string templatePath = _settings.LoadSettings().ExcelTemplatePath;
        IReadOnlyList<GroupTraceAggregate> aggregates = _trace.RecalculateByGroups();
        _export.ToExcelInput(templatePath, aggregates.Select(ToExcelInputRow).ToList());
        IReadOnlyList<ExcelOutputRow> outputRows = _export.GetCachedOrLoadOutput(templatePath);
        if (outputRows.Count > 0)
        {
            _export.ExportExcelOutputReportCsv(outputRows);
            _export.ToAutoCadTableFromOutput(outputRows, new Point3d(0, -120, 0));
        }

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
        IReadOnlyList<GroupTraceAggregate> aggregates = _trace.RecalculateByGroups();
        _log.Write($"EOM_ОБНОВИТЬ завершена. Групп в расчете: {aggregates.Count}.");
        // END_BLOCK_COMMAND_EOM_UPDATE
    }

    [CommandMethod(PluginConfig.Commands.ExportExcel)]
    public void EomExportExcel()
    {
        // START_BLOCK_COMMAND_EOM_EXPORT_EXCEL
        IReadOnlyList<GroupTraceAggregate> aggregates = _trace.RecalculateByGroups();
        string templatePath = _settings.LoadSettings().ExcelTemplatePath;
        _export.ClearCachedOutput(templatePath);
        var inputRows = aggregates.Select(ToExcelInputRow).ToList();
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
        string templatePath = _settings.LoadSettings().ExcelTemplatePath;
        bool cacheHit = _export.TryGetCachedOutput(templatePath, out IReadOnlyList<ExcelOutputRow> rows);
        if (!cacheHit)
        {
            rows = _export.GetCachedOrLoadOutput(templatePath);
        }

        if (rows.Count == 0)
        {
            _log.Write("EOM_ПОСТРОИТЬ_ОЛС: нет строк OUTPUT для построения.");
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
        string templatePath = _settings.LoadSettings().ExcelTemplatePath;
        bool cacheHit = _export.TryGetCachedOutput(templatePath, out IReadOnlyList<ExcelOutputRow> rows);
        if (!cacheHit)
        {
            rows = _export.GetCachedOrLoadOutput(templatePath);
        }

        if (rows.Count == 0)
        {
            _log.Write("EOM_КОМПОНОВКА_ЩИТА: нет строк OUTPUT для построения.");
            return;
        }

        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PromptPointResult pointResult = doc.Editor.GetPoint("\nУкажите точку вставки компоновки щита: ");
        if (pointResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_КОМПОНОВКА_ЩИТА отменена.");
            return;
        }

        int modulesPerRow = _settings.LoadSettings().PanelModulesPerRow;
        DrawPanelLayout(rows, pointResult.Value, modulesPerRow);
        _log.Write($"EOM_КОМПОНОВКА_ЩИТА: строк {rows.Count}, модулей в ряду {modulesPerRow}, источник={(cacheHit ? "кэш" : "файл")}.");
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
                        continue;
                    }

                    string groupValue = group.Trim();
                    loadGroups.Add(groupValue);
                    if (groupRegex is not null && !groupRegex.IsMatch(groupValue))
                    {
                        regexMismatchGroups++;
                    }

                    if (string.IsNullOrWhiteSpace(power) || !double.TryParse(power.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
                    {
                        invalidPowerLoads++;
                    }

                    if (!shieldsByGroup.TryGetValue(groupValue, out HashSet<string>? shields))
                    {
                        shields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        shieldsByGroup[groupValue] = shields;
                    }

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

        _log.Write(
            $"EOM_ПРОВЕРКА: линии без {PluginConfig.Strings.Group}={missingGroupLines}; " +
            $"нагрузки без {PluginConfig.Strings.Group}={missingGroupLoads}; " +
            $"{PluginConfig.Strings.Power} нечисловая={invalidPowerLoads}; " +
            $"группы нагрузок без линий={groupsWithoutLines}; " +
            $"линии с типом по умолчанию={defaultInstallTypeLines}; " +
            $"несовпадение {PluginConfig.Strings.Shield} внутри группы={groupShieldMismatch}; " +
            $"regex-ошибки группы={regexMismatchGroups}.");
        // END_BLOCK_COMMAND_EOM_VALIDATE
    }

    private void OnObjectAppended(ObjectId objectId)
    {
        // START_BLOCK_ON_OBJECT_APPENDED
        _xdata.ApplyActiveGroupToEntity(objectId);
        // END_BLOCK_ON_OBJECT_APPENDED
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

                y -= step;
            }

            tr.Commit();
        }
        // END_BLOCK_DRAW_OLS_ROWS
    }

    private static void DrawPanelLayout(IReadOnlyList<ExcelOutputRow> rows, Point3d basePoint, int modulesPerRow)
    {
        // START_BLOCK_DRAW_PANEL_LAYOUT
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null || rows.Count == 0)
        {
            return;
        }

        int safeModulesPerRow = modulesPerRow <= 0 ? 24 : modulesPerRow;
        const double moduleWidth = 5.0;
        const double moduleHeight = 4.0;
        const double rowGap = 2.0;

        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            double currentX = basePoint.X;
            double currentY = basePoint.Y;
            int occupiedInRow = 0;

            foreach (IGrouping<string, ExcelOutputRow> shieldGroup in rows
                .OrderBy(x => x.Shield)
                .ThenBy(x => x.Group)
                .GroupBy(x => x.Shield, StringComparer.OrdinalIgnoreCase))
            {
                occupiedInRow = 0;
                currentX = basePoint.X;
                currentY -= 8.0;

                var shieldTitle = new DBText
                {
                    Position = new Point3d(currentX, currentY + 3.5, basePoint.Z),
                    Height = 2.8,
                    TextString = $"{PluginConfig.Strings.Shield}: {shieldGroup.Key}",
                    Layer = "0"
                };
                ms.AppendEntity(shieldTitle);
                tr.AddNewlyCreatedDBObject(shieldTitle, true);

                var inputDevice = new LayoutDevice("\u0412\u0412\u041e\u0414", 2);
                PlaceLayoutDevice(ms, tr, inputDevice, basePoint.Z, basePoint.X, ref currentX, ref currentY, ref occupiedInRow, safeModulesPerRow, moduleWidth, moduleHeight, rowGap);

                foreach (ExcelOutputRow row in shieldGroup)
                {
                    if (row.RcdModules > 0)
                    {
                        var rcdDevice = new LayoutDevice($"{row.Group}: {row.RcdDiff}", Math.Max(1, row.RcdModules));
                        PlaceLayoutDevice(ms, tr, rcdDevice, basePoint.Z, basePoint.X, ref currentX, ref currentY, ref occupiedInRow, safeModulesPerRow, moduleWidth, moduleHeight, rowGap);
                    }

                    var breakerDevice = new LayoutDevice($"{row.Group}: {row.CircuitBreaker}", Math.Max(1, row.CircuitBreakerModules));
                    PlaceLayoutDevice(ms, tr, breakerDevice, basePoint.Z, basePoint.X, ref currentX, ref currentY, ref occupiedInRow, safeModulesPerRow, moduleWidth, moduleHeight, rowGap);
                }

                currentY -= moduleHeight + rowGap + 8.0;
            }

            tr.Commit();
        }
        // END_BLOCK_DRAW_PANEL_LAYOUT
    }

    private static void PlaceLayoutDevice(
        BlockTableRecord ms,
        Transaction tr,
        LayoutDevice device,
        double z,
        double rowStartX,
        ref double currentX,
        ref double currentY,
        ref int occupiedInRow,
        int modulesPerRow,
        double moduleWidth,
        double moduleHeight,
        double rowGap)
    {
        // START_BLOCK_PLACE_LAYOUT_DEVICE
        if (occupiedInRow + device.ModuleCount > modulesPerRow)
        {
            occupiedInRow = 0;
            currentX = rowStartX;
            currentY -= moduleHeight + rowGap + 4.0;
        }

        var frame = new Polyline();
        frame.AddVertexAt(0, new Point2d(currentX, currentY), 0, 0, 0);
        frame.AddVertexAt(1, new Point2d(currentX + moduleWidth * device.ModuleCount, currentY), 0, 0, 0);
        frame.AddVertexAt(2, new Point2d(currentX + moduleWidth * device.ModuleCount, currentY - moduleHeight), 0, 0, 0);
        frame.AddVertexAt(3, new Point2d(currentX, currentY - moduleHeight), 0, 0, 0);
        frame.Closed = true;
        ms.AppendEntity(frame);
        tr.AddNewlyCreatedDBObject(frame, true);

        var label = new DBText
        {
            Position = new Point3d(currentX, currentY + 1.0, z),
            Height = 2.0,
            TextString = $"{device.Label} [{device.ModuleCount}]",
            Layer = "0"
        };
        ms.AppendEntity(label);
        tr.AddNewlyCreatedDBObject(label, true);

        currentX += moduleWidth * device.ModuleCount;
        occupiedInRow += device.ModuleCount;
        // END_BLOCK_PLACE_LAYOUT_DEVICE
    }

    private sealed record LayoutDevice(string Label, int ModuleCount);

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
}
