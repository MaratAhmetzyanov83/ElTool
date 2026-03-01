// FILE: src/Commands/CommandRegistry.cs
// VERSION: 1.7.5
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
//   EomExcelPath - Sets explicit path to workbook used by Excel commands.
//   EomBuildOls - Draws one-line diagram from Excel OUTPUT.
//   EomPanelLayoutConfig - Opens UI editor for panel layout map and SOURCE->LAYOUT bindings.
//   EomBuildPanelLayout - Builds panel layout from user-selected OLS blocks using panel mapping configuration.
//   EomBindPanelLayoutVisualization - Creates and stores source OLS to visualization block selector rule.
//   BuildPanelLayoutModel - Converts mapped OLS devices to normalized DIN placement model.
// END_MODULE_MAP
//
// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.7.5 - EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ now can also insert an AcExcel-linked table (default yes) after scheme generation, so table view parity and scheme output are available in one flow.
// END_CHANGE_SUMMARY

using Autodesk.AutoCAD.Runtime;
using System.Text;
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
    // START_CONTRACT: EomMap
    //   PURPOSE: Eom map.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomMap

    public void EomMap()
    {
        // START_BLOCK_COMMAND_EOM_MAP
        if (_license.Validate() is LicenseState.Invalid or LicenseState.Expired)
        {
            _log.Write("Р С™Р С•Р СР В°Р Р…Р Т‘Р В° EOM_MAP Р В·Р В°Р В±Р В»Р С•Р С”Р С‘РЎР‚Р С•Р Р†Р В°Р Р…Р В° Р В»Р С‘РЎвЂ Р ВµР Р…Р В·Р С‘Р ВµР в„–.");
            return;
        }

        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            _log.Write("Р С’Р С”РЎвЂљР С‘Р Р†Р Р…РЎвЂ№Р в„– Р Т‘Р С•Р С”РЎС“Р СР ВµР Р…РЎвЂљ Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р….");
            return;
        }

        Editor editor = doc.Editor;
        var sourceOptions = new PromptEntityOptions("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ Р С‘РЎРѓРЎвЂ¦Р С•Р Т‘Р Р…РЎвЂ№Р в„– Р В±Р В»Р С•Р С” Р Т‘Р В»РЎРЏ Р В·Р В°Р СР ВµР Р…РЎвЂ№: ");
        sourceOptions.SetRejectMessage("\nР СњРЎС“Р В¶Р ВµР Р… Р В±Р В»Р С•Р С” (BlockReference).");
        sourceOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult sourceResult = editor.GetEntity(sourceOptions);
        if (sourceResult.Status != PromptStatus.OK)
        {
            _log.Write("Р С™Р С•Р СР В°Р Р…Р Т‘Р В° Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°: Р С‘РЎРѓРЎвЂ¦Р С•Р Т‘Р Р…РЎвЂ№Р в„– Р В±Р В»Р С•Р С” Р Р…Р Вµ Р Р†РЎвЂ№Р В±РЎР‚Р В°Р Р….");
            return;
        }

        var targetOptions = new PromptEntityOptions("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ РЎвЂ Р ВµР В»Р ВµР Р†Р С•Р в„– Р В±Р В»Р С•Р С”: ");
        targetOptions.SetRejectMessage("\nР СњРЎС“Р В¶Р ВµР Р… Р В±Р В»Р С•Р С” (BlockReference).");
        targetOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult targetResult = editor.GetEntity(targetOptions);
        if (targetResult.Status != PromptStatus.OK)
        {
            _log.Write("Р С™Р С•Р СР В°Р Р…Р Т‘Р В° Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°: РЎвЂ Р ВµР В»Р ВµР Р†Р С•Р в„– Р В±Р В»Р С•Р С” Р Р…Р Вµ Р Р†РЎвЂ№Р В±РЎР‚Р В°Р Р….");
            return;
        }

        int replaced = _mapping.ExecuteMapping(sourceResult.ObjectId, targetResult.ObjectId);
        _log.Write($"EOM_MAP Р В·Р В°Р Р†Р ВµРЎР‚РЎв‚¬Р ВµР Р…Р В°. Р вЂ”Р В°Р СР ВµР Р…Р ВµР Р…Р С• Р В±Р В»Р С•Р С”Р С•Р Р†: {replaced}.");
        // END_BLOCK_COMMAND_EOM_MAP
    }

    [CommandMethod("EOM_TRACE")]
    // START_CONTRACT: EomTrace
    //   PURPOSE: Eom trace.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomTrace

    public void EomTrace()
    {
        // START_BLOCK_COMMAND_EOM_TRACE
        if (_license.Validate() is LicenseState.Invalid or LicenseState.Expired)
        {
            _log.Write("Р С™Р С•Р СР В°Р Р…Р Т‘Р В° EOM_TRACE Р В·Р В°Р В±Р В»Р С•Р С”Р С‘РЎР‚Р С•Р Р†Р В°Р Р…Р В° Р В»Р С‘РЎвЂ Р ВµР Р…Р В·Р С‘Р ВµР в„–.");
            return;
        }

        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        Editor editor = doc.Editor;

        var baseLineOptions = new PromptEntityOptions("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ Р В±Р В°Р В·Р С•Р Р†РЎС“РЎР‹ Р С—Р С•Р В»Р С‘Р В»Р С‘Р Р…Р С‘РЎР‹ (Р СР В°Р С–Р С‘РЎРѓРЎвЂљРЎР‚Р В°Р В»РЎРЉ): ");
        baseLineOptions.SetRejectMessage("\nР СњРЎС“Р В¶Р Р…Р В° Р С—Р С•Р В»Р С‘Р В»Р С‘Р Р…Р С‘РЎРЏ.");
        baseLineOptions.AddAllowedClass(typeof(Polyline), true);
        PromptEntityResult baseLineResult = editor.GetEntity(baseLineOptions);
        if (baseLineResult.Status != PromptStatus.OK)
        {
            _log.Write("Р С™Р С•Р СР В°Р Р…Р Т‘Р В° Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°: Р В±Р В°Р В·Р С•Р Р†Р В°РЎРЏ Р С—Р С•Р В»Р С‘Р В»Р С‘Р Р…Р С‘РЎРЏ Р Р…Р Вµ Р Р†РЎвЂ№Р В±РЎР‚Р В°Р Р…Р В°.");
            return;
        }

        int created = 0;
        while (true)
        {
            var targetBlockOptions = new PromptEntityOptions("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ Р В±Р В»Р С•Р С” Р Т‘Р В»РЎРЏ Р С•РЎвЂљР Р†Р ВµРЎвЂљР Р†Р В»Р ВµР Р…Р С‘РЎРЏ [Enter - Р В·Р В°Р Р†Р ВµРЎР‚РЎв‚¬Р С‘РЎвЂљРЎРЉ]: ")
            {
                AllowNone = true
            };
            targetBlockOptions.SetRejectMessage("\nР СњРЎС“Р В¶Р ВµР Р… Р В±Р В»Р С•Р С” (BlockReference).");
            targetBlockOptions.AddAllowedClass(typeof(BlockReference), true);
            PromptEntityResult targetBlockResult = editor.GetEntity(targetBlockOptions);

            if (targetBlockResult.Status == PromptStatus.None)
            {
                _log.Write($"EOM_TRACE Р В·Р В°Р Р†Р ВµРЎР‚РЎв‚¬Р ВµР Р…Р В°. Р СџР С•РЎРѓРЎвЂљРЎР‚Р С•Р ВµР Р…Р С• Р С•РЎвЂљР Р†Р ВµРЎвЂљР Р†Р В»Р ВµР Р…Р С‘Р в„–: {created}.");
                break;
            }

            if (targetBlockResult.Status != PromptStatus.OK)
            {
                _log.Write("Р С™Р С•Р СР В°Р Р…Р Т‘Р В° Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В° Р С—Р С•Р В»РЎРЉР В·Р С•Р Р†Р В°РЎвЂљР ВµР В»Р ВµР С.");
                break;
            }

            TraceResult? trace = _trace.ExecuteTraceFromBase(baseLineResult.ObjectId, targetBlockResult.ObjectId);
            if (trace is null)
            {
                _log.Write("Р СњР Вµ РЎС“Р Т‘Р В°Р В»Р С•РЎРѓРЎРЉ Р С—Р С•РЎРѓРЎвЂљРЎР‚Р С•Р С‘РЎвЂљРЎРЉ Р С•РЎвЂљР Р†Р ВµРЎвЂљР Р†Р В»Р ВµР Р…Р С‘Р Вµ Р Т‘Р В»РЎРЏ Р Р†РЎвЂ№Р В±РЎР‚Р В°Р Р…Р Р…Р С•Р С–Р С• Р В±Р В»Р С•Р С”Р В°.");
                continue;
            }

            created++;
            _log.Write($"Р С›РЎвЂљР Р†Р ВµРЎвЂљР Р†Р В»Р ВµР Р…Р С‘Р Вµ #{created}: {trace.TotalLength:0.###} Р С.");
        }
        // END_BLOCK_COMMAND_EOM_TRACE
    }

    [CommandMethod("EOM_SPEC")]
    // START_CONTRACT: EomSpec
    //   PURPOSE: Eom spec.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomSpec

    public void EomSpec()
    {
        // START_BLOCK_COMMAND_EOM_SPEC
        if (_license.Validate() is LicenseState.Invalid or LicenseState.Expired)
        {
            _log.Write("Р С™Р С•Р СР В°Р Р…Р Т‘Р В° EOM_SPEC Р В·Р В°Р В±Р В»Р С•Р С”Р С‘РЎР‚Р С•Р Р†Р В°Р Р…Р В° Р В»Р С‘РЎвЂ Р ВµР Р…Р В·Р С‘Р ВµР в„–.");
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
            bool outputCsvExists = File.Exists(outputPath);
            bool workbookExists = File.Exists(templatePath);
            if (!outputCsvExists && !workbookExists)
            {
                _log.Write($"EOM_SPEC: Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р… OUTPUT ({outputPath}) Р С‘ РЎв‚¬Р В°Р В±Р В»Р С•Р Р… Excel ({templatePath}).");
            }
            else
            {
                if (!outputCsvExists && workbookExists)
                {
                    _log.Write($"EOM_SPEC: OUTPUT Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р… ({outputPath}), Р С‘РЎРѓР С—Р С•Р В»РЎРЉР В·РЎС“Р ВµРЎвЂљРЎРѓРЎРЏ fallback Р С‘Р В· Р С”Р Р…Р С‘Р С–Р С‘ '{templatePath}' (Р В»Р С‘РЎРѓРЎвЂљ 'Р вЂ™ Р С’Р С”Р В°Р Т‘').");
                }

                outputRows = _export.GetCachedOrLoadOutput(templatePath);
            }
        }

        if (outputRows.Count > 0)
        {
            _export.ExportExcelOutputReportCsv(outputRows);
            _export.ToAutoCadTableFromOutput(outputRows, new Point3d(0, -120, 0));
        }

        string inputPath = GetExpectedExcelInputPath(templatePath);
        _log.Write($"EOM_SPEC Р В·Р В°Р Р†Р ВµРЎР‚РЎв‚¬Р ВµР Р…Р В°. INPUT: {inputPath}");
        // END_BLOCK_COMMAND_EOM_SPEC
    }

    [CommandMethod(PluginConfig.Commands.MapConfig)]
    // START_CONTRACT: EomMapCfg
    //   PURPOSE: Eom map cfg.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomMapCfg

    public void EomMapCfg()
    {
        // START_BLOCK_COMMAND_EOM_MAPCFG
        try
        {
            var vm = new MappingConfigWindowViewModel();
            var window = new MappingConfigWindow(vm);
            Application.ShowModalWindow(window);
            _log.Write("Р С›Р С”Р Р…Р С• Р Р…Р В°РЎРѓРЎвЂљРЎР‚Р С•Р в„–Р С”Р С‘ РЎРѓР С•Р С•РЎвЂљР Р†Р ВµРЎвЂљРЎРѓРЎвЂљР Р†Р С‘Р в„– Р В·Р В°Р С”РЎР‚РЎвЂ№РЎвЂљР С•.");
        }
        catch (System.Exception ex)
        {
            _log.Write($"Р С›РЎв‚¬Р С‘Р В±Р С”Р В° Р С•РЎвЂљР С”РЎР‚РЎвЂ№РЎвЂљР С‘РЎРЏ Р С•Р С”Р Р…Р В° Р Р…Р В°РЎРѓРЎвЂљРЎР‚Р С•Р в„–Р С”Р С‘ РЎРѓР С•Р С•РЎвЂљР Р†Р ВµРЎвЂљРЎРѓРЎвЂљР Р†Р С‘Р в„–: {ex.Message}");
        }
        // END_BLOCK_COMMAND_EOM_MAPCFG
    }

    [CommandMethod(PluginConfig.Commands.PanelLayoutConfig)]
    // START_CONTRACT: EomPanelLayoutConfig
    //   PURPOSE: Eom panel layout config.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomPanelLayoutConfig

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
            _log.Write("Р С›Р С”Р Р…Р С• Р Р…Р В°РЎРѓРЎвЂљРЎР‚Р С•Р в„–Р С”Р С‘ Р С”Р С•Р СР С—Р С•Р Р…Р С•Р Р†Р С”Р С‘ РЎвЂ°Р С‘РЎвЂљР В° Р В·Р В°Р С”РЎР‚РЎвЂ№РЎвЂљР С•.");
        }
        catch (System.Exception ex)
        {
            _log.Write($"Р С›РЎв‚¬Р С‘Р В±Р С”Р В° Р С•РЎвЂљР С”РЎР‚РЎвЂ№РЎвЂљР С‘РЎРЏ Р С•Р С”Р Р…Р В° Р Р…Р В°РЎРѓРЎвЂљРЎР‚Р С•Р в„–Р С”Р С‘ Р С”Р С•Р СР С—Р С•Р Р…Р С•Р Р†Р С”Р С‘ РЎвЂ°Р С‘РЎвЂљР В°: {ex.Message}");
        }
        // END_BLOCK_COMMAND_EOM_PANEL_LAYOUT_CONFIG
    }

    [CommandMethod(PluginConfig.Commands.ActiveGroup)]
    // START_CONTRACT: EomActiveGroup
    //   PURPOSE: Eom active group.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomActiveGroup

    public void EomActiveGroup()
    {
        // START_BLOCK_COMMAND_EOM_ACTIVE_GROUP
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PromptResult prompt = doc.Editor.GetString("\nР вЂ™Р Р†Р ВµР Т‘Р С‘РЎвЂљР Вµ Р С”Р С•Р Т‘ Р В°Р С”РЎвЂљР С‘Р Р†Р Р…Р С•Р в„– Р С–РЎР‚РЎС“Р С—Р С—РЎвЂ№: ");
        if (prompt.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(prompt.StringResult))
        {
            _log.Write("Р С’Р С”РЎвЂљР С‘Р Р†Р Р…Р В°РЎРЏ Р С–РЎР‚РЎС“Р С—Р С—Р В° Р Р…Р Вµ РЎС“РЎРѓРЎвЂљР В°Р Р…Р С•Р Р†Р В»Р ВµР Р…Р В°.");
            return;
        }

        _xdata.SetActiveGroup(prompt.StringResult.Trim());
        _acad.SubscribeObjectAppended(OnObjectAppended);
        _log.Write($"Р С’Р С”РЎвЂљР С‘Р Р†Р Р…Р В°РЎРЏ Р С–РЎР‚РЎС“Р С—Р С—Р В° РЎС“РЎРѓРЎвЂљР В°Р Р…Р С•Р Р†Р В»Р ВµР Р…Р В°: {prompt.StringResult.Trim()}");
        // END_BLOCK_COMMAND_EOM_ACTIVE_GROUP
    }

    [CommandMethod(PluginConfig.Commands.AssignGroup)]
    // START_CONTRACT: EomAssignGroup
    //   PURPOSE: Eom assign group.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomAssignGroup

    public void EomAssignGroup()
    {
        // START_BLOCK_COMMAND_EOM_ASSIGN_GROUP
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PromptResult groupPrompt = doc.Editor.GetString("\nР вЂ™Р Р†Р ВµР Т‘Р С‘РЎвЂљР Вµ Р С”Р С•Р Т‘ Р С–РЎР‚РЎС“Р С—Р С—РЎвЂ№: ");
        if (groupPrompt.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(groupPrompt.StringResult))
        {
            _log.Write("Р вЂњРЎР‚РЎС“Р С—Р С—Р В° Р Р…Р Вµ Р В·Р В°Р Т‘Р В°Р Р…Р В°.");
            return;
        }

        var selectionOptions = new PromptSelectionOptions { MessageForAdding = "\nР вЂ™РЎвЂ№Р В±Р ВµРЎР‚Р С‘РЎвЂљР Вµ Р В»Р С‘Р Р…Р С‘Р С‘/Р С—Р С•Р В»Р С‘Р В»Р С‘Р Р…Р С‘Р С‘: " };
        var filter = new SelectionFilter([
            new TypedValue((int)DxfCode.Start, "LINE,LWPOLYLINE")
        ]);
        PromptSelectionResult selection = doc.Editor.GetSelection(selectionOptions, filter);
        if (selection.Status != PromptStatus.OK)
        {
            _log.Write("Р СњР В°Р В·Р Р…Р В°РЎвЂЎР ВµР Р…Р С‘Р Вµ Р С–РЎР‚РЎС“Р С—Р С—РЎвЂ№ Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р С•.");
            return;
        }

        _xdata.AssignGroupToSelection(selection.Value.GetObjectIds(), groupPrompt.StringResult.Trim());
        _log.Write($"Р вЂњРЎР‚РЎС“Р С—Р С—Р В° {groupPrompt.StringResult.Trim()} Р Р…Р В°Р В·Р Р…Р В°РЎвЂЎР ВµР Р…Р В°: {selection.Value.Count} Р С•Р В±РЎР‰Р ВµР С”РЎвЂљР С•Р Р†.");
        // END_BLOCK_COMMAND_EOM_ASSIGN_GROUP
    }

    [CommandMethod(PluginConfig.Commands.InstallTypeSettings)]
    // START_CONTRACT: EomInstallTypeSettings
    //   PURPOSE: Eom install type settings.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomInstallTypeSettings

    public void EomInstallTypeSettings()
    {
        // START_BLOCK_COMMAND_EOM_INSTALL_TYPE_SETTINGS
        string path = _settings.OpenInstallTypeConfig();
        _log.Write($"Р С™Р С•Р Р…РЎвЂћР С‘Р С– Р С—РЎР‚Р В°Р Р†Р С‘Р В» Р С—РЎР‚Р С•Р С”Р В»Р В°Р Т‘Р С”Р С‘: {path}");
        // END_BLOCK_COMMAND_EOM_INSTALL_TYPE_SETTINGS
    }

    [CommandMethod(PluginConfig.Commands.ExcelPath)]
    // START_CONTRACT: EomExcelPath
    //   PURPOSE: Eom excel path.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomExcelPath

    public void EomExcelPath()
    {
        // START_BLOCK_COMMAND_EOM_EXCEL_PATH
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        var options = new PromptStringOptions("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ Р С—Р С•Р В»Р Р…РЎвЂ№Р в„– Р С—РЎС“РЎвЂљРЎРЉ Р С” Excel (*.xlsx/*.xlsm): ")
        {
            AllowSpaces = true
        };
        PromptResult prompt = doc.Editor.GetString(options);
        if (prompt.Status != PromptStatus.OK)
        {
            _log.Write("Р СџРЎС“РЎвЂљРЎРЉ Excel Р Р…Р Вµ Р С‘Р В·Р СР ВµР Р…Р ВµР Р….");
            return;
        }

        string rawPath = (prompt.StringResult ?? string.Empty).Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(rawPath))
        {
            _log.Write("Р СџРЎС“РЎвЂљРЎРЉ Excel Р Р…Р Вµ Р В·Р В°Р Т‘Р В°Р Р….");
            return;
        }

        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(rawPath);
        }
        catch
        {
            _log.Write($"Р СњР ВµР С”Р С•РЎР‚РЎР‚Р ВµР С”РЎвЂљР Р…РЎвЂ№Р в„– Р С—РЎС“РЎвЂљРЎРЉ Excel: {rawPath}");
            return;
        }

        string extension = Path.GetExtension(fullPath);
        bool isWorkbook = extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".xltx", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".xltm", StringComparison.OrdinalIgnoreCase);
        if (!isWorkbook)
        {
            _log.Write($"Р СњР ВµР С—Р С•Р Т‘Р Т‘Р ВµРЎР‚Р В¶Р С‘Р Р†Р В°Р ВµР СРЎвЂ№Р в„– РЎвЂћР С•РЎР‚Р СР В°РЎвЂљ: {extension}. Р Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ РЎвЂћР В°Р в„–Р В» .xlsx/.xlsm.");
            return;
        }

        if (!File.Exists(fullPath))
        {
            _log.Write($"Р В¤Р В°Р в„–Р В» Excel Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р…: {fullPath}");
            return;
        }

        SettingsModel settings = _settings.LoadSettings();
        settings.ExcelTemplatePath = fullPath;
        _settings.SaveSettings(settings);
        _export.ClearCachedOutput(fullPath);
        _log.Write($"Р СџРЎС“РЎвЂљРЎРЉ Excel РЎРѓР С•РЎвЂ¦РЎР‚Р В°Р Р…Р ВµР Р…: {fullPath}");
        // END_BLOCK_COMMAND_EOM_EXCEL_PATH
    }

    [CommandMethod(PluginConfig.Commands.Update)]
    // START_CONTRACT: EomUpdate
    //   PURPOSE: Eom update.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomUpdate

    public void EomUpdate()
    {
        // START_BLOCK_COMMAND_EOM_UPDATE
        IReadOnlyList<GroupTraceAggregate> aggregates = _trace.RecalculateByGroups();
        _log.Write($"EOM_Р С›Р вЂР СњР С›Р вЂ™Р ВР СћР В¬ Р В·Р В°Р Р†Р ВµРЎР‚РЎв‚¬Р ВµР Р…Р В°. Р вЂњРЎР‚РЎС“Р С—Р С— Р Р† РЎР‚Р В°РЎРѓРЎвЂЎР ВµРЎвЂљР Вµ: {aggregates.Count}.");
        // END_BLOCK_COMMAND_EOM_UPDATE
    }

    [CommandMethod(PluginConfig.Commands.ExportExcel)]
    // START_CONTRACT: EomExportExcel
    //   PURPOSE: Eom export excel.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomExportExcel

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
        _log.Write($"EOM_Р В­Р С™Р РЋР СџР С›Р В Р Сћ_EXCEL Р В·Р В°Р Р†Р ВµРЎР‚РЎв‚¬Р ВµР Р…Р В°. INPUT: {GetExpectedExcelInputPath(templatePath)}");
        // END_BLOCK_COMMAND_EOM_EXPORT_EXCEL
    }

    [CommandMethod(PluginConfig.Commands.ImportExcel)]
    // START_CONTRACT: EomImportExcel
    //   PURPOSE: Eom import excel.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomImportExcel

    public void EomImportExcel()
    {
        // START_BLOCK_COMMAND_EOM_IMPORT_EXCEL
        if (!TryResolveExcelTemplatePath(PluginConfig.Commands.ImportExcel, out string templatePath))
        {
            return;
        }

        string outputPath = GetExpectedExcelOutputPath(templatePath);
        bool outputCsvExists = File.Exists(outputPath);
        bool workbookExists = File.Exists(templatePath);
        _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL: template='{templatePath}', output='{outputPath}', output_exists={outputCsvExists}, workbook_exists={workbookExists}.");
        if (!outputCsvExists && !workbookExists)
        {
            _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL: Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р… OUTPUT ({outputPath}) Р С‘ РЎв‚¬Р В°Р В±Р В»Р С•Р Р… Excel ({templatePath}).");
            return;
        }

        if (outputCsvExists)
        {
            _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL: Р С‘РЎРѓРЎвЂљР С•РЎвЂЎР Р…Р С‘Р С” Р Т‘Р В°Р Р…Р Р…РЎвЂ№РЎвЂ¦ OUTPUT.csv ({outputPath}).");
        }
        else if (workbookExists)
        {
            _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL: OUTPUT Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р… ({outputPath}), Р С‘РЎРѓР С—Р С•Р В»РЎРЉР В·РЎС“Р ВµРЎвЂљРЎРѓРЎРЏ fallback Р С‘Р В· Р С”Р Р…Р С‘Р С–Р С‘ '{templatePath}' (Р В»Р С‘РЎРѓРЎвЂљ 'Р вЂ™ Р С’Р С”Р В°Р Т‘').");
        }

        IReadOnlyList<ExcelOutputRow> rows = _export.FromExcelOutput(templatePath);
        if (rows.Count > 0)
        {
            _export.ExportExcelOutputReportCsv(rows);
        }
        else
        {
            _log.Write("EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL: Р С‘Р СР С—Р С•РЎР‚РЎвЂљ Р Р†РЎвЂ№Р С—Р С•Р В»Р Р…Р ВµР Р…, Р Р…Р С• Р Р†Р В°Р В»Р С‘Р Т‘Р Р…РЎвЂ№Р Вµ РЎРѓРЎвЂљРЎР‚Р С•Р С”Р С‘ Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р…РЎвЂ№.");
        }

        if (workbookExists)
        {
            Document? doc = Application.DocumentManager.MdiActiveDocument;
            if (doc is not null)
            {
                var tablePrompt = new PromptKeywordOptions("\nР РЋР С•Р В·Р Т‘Р В°РЎвЂљРЎРЉ РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…РЎС“РЎР‹ РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎС“ Excel [Р вЂќР В°/Р СњР ВµРЎвЂљ] <Р вЂќР В°>: ");
                tablePrompt.Keywords.Add("Р вЂќР В°");
                tablePrompt.Keywords.Add("Р СњР ВµРЎвЂљ");
                tablePrompt.AllowNone = true;
                PromptResult tableDecision = doc.Editor.GetKeywords(tablePrompt);
                bool shouldInsertTable = tableDecision.Status switch
                {
                    PromptStatus.OK => !string.Equals(tableDecision.StringResult, "Р СњР ВµРЎвЂљ", StringComparison.OrdinalIgnoreCase),
                    PromptStatus.None => true,
                    _ => false
                };

                if (shouldInsertTable)
                {
                    PromptPointResult tablePoint = doc.Editor.GetPoint("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ РЎвЂљР С•РЎвЂЎР С”РЎС“ Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р С‘ РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…Р С•Р в„– РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎвЂ№ Excel: ");
                    if (tablePoint.Status == PromptStatus.OK)
                    {
                        bool tableInserted = _export.TryInsertExcelLinkedTable(templatePath, tablePoint.Value, out string tableStatus);
                        _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL: {tableStatus}");
                        if (!tableInserted)
                        {
                            _log.Write("EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL: Р С›Р вЂєР РЋ Р С‘Р СР С—Р С•РЎР‚РЎвЂљ/Р С”РЎРЊРЎв‚¬ РЎРѓР С•РЎвЂ¦РЎР‚Р В°Р Р…Р ВµР Р…РЎвЂ№, Р Р…Р С• Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р В° РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…Р С•Р в„– РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎвЂ№ Р Р…Р Вµ Р Р†РЎвЂ№Р С—Р С•Р В»Р Р…Р ВµР Р…Р В°.");
                        }
                    }
                    else
                    {
                        _log.Write("EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL: Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р В° РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…Р С•Р в„– РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎвЂ№ Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В° Р С—Р С•Р В»РЎРЉР В·Р С•Р Р†Р В°РЎвЂљР ВµР В»Р ВµР С.");
                    }
                }
            }
        }

        _log.Write($"EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL Р В·Р В°Р Р†Р ВµРЎР‚РЎв‚¬Р ВµР Р…Р В°. Р ВР СР С—Р С•РЎР‚РЎвЂљР С‘РЎР‚Р С•Р Р†Р В°Р Р…Р С• РЎРѓРЎвЂљРЎР‚Р С•Р С”: {rows.Count}. Р С™РЎРЊРЎв‚¬ Р С•Р В±Р Р…Р С•Р Р†Р В»Р ВµР Р….");
        // END_BLOCK_COMMAND_EOM_IMPORT_EXCEL
    }

    [CommandMethod(PluginConfig.Commands.BuildOls)]
    // START_CONTRACT: EomBuildOls
    //   PURPOSE: Eom build ols.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomBuildOls

    public void EomBuildOls()
    {
        // START_BLOCK_COMMAND_EOM_BUILD_OLS
        if (!TryResolveExcelTemplatePath(PluginConfig.Commands.BuildOls, out string templatePath))
        {
            return;
        }

        string outputPath = GetExpectedExcelOutputPath(templatePath);
        bool outputCsvExists = File.Exists(outputPath);
        bool workbookExists = File.Exists(templatePath);
        bool cacheHit = _export.TryGetCachedOutput(templatePath, out IReadOnlyList<ExcelOutputRow> rows);
        if (!cacheHit)
        {
            if (!outputCsvExists && !workbookExists)
            {
                _log.Write($"EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р… OUTPUT ({outputPath}) Р С‘ РЎв‚¬Р В°Р В±Р В»Р С•Р Р… Excel ({templatePath}). Р вЂ™РЎвЂ№Р С—Р С•Р В»Р Р…Р С‘РЎвЂљР Вµ EOM_Р ВР СљР СџР С›Р В Р Сћ_EXCEL.");
                return;
            }

            if (!outputCsvExists && workbookExists)
            {
                _log.Write($"EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: OUTPUT Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р… ({outputPath}), Р С‘РЎРѓР С—Р С•Р В»РЎРЉР В·РЎС“Р ВµРЎвЂљРЎРѓРЎРЏ fallback Р С‘Р В· Р С”Р Р…Р С‘Р С–Р С‘ '{templatePath}' (Р В»Р С‘РЎРѓРЎвЂљ 'Р вЂ™ Р С’Р С”Р В°Р Т‘').");
            }

            rows = _export.GetCachedOrLoadOutput(templatePath);
        }

        if (rows.Count == 0)
        {
            _log.Write(outputCsvExists
                ? $"EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: OUTPUT Р С—РЎС“РЎРѓРЎвЂљР С•Р в„– ({outputPath})."
                : $"EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: Р Р† Р С”Р Р…Р С‘Р С–Р Вµ '{templatePath}' (Р В»Р С‘РЎРѓРЎвЂљ 'Р вЂ™ Р С’Р С”Р В°Р Т‘') Р Р…Р ВµРЎвЂљ Р Р†Р В°Р В»Р С‘Р Т‘Р Р…РЎвЂ№РЎвЂ¦ РЎРѓРЎвЂљРЎР‚Р С•Р С”.");
            return;
        }

        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PromptPointResult pointResult = doc.Editor.GetPoint("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ РЎвЂљР С•РЎвЂЎР С”РЎС“ Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р С‘ Р С›Р вЂєР РЋ: ");
        if (pointResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°.");
            return;
        }

        PromptResult shieldResult = doc.Editor.GetString("\nР В©Р ВР Сћ Р Т‘Р В»РЎРЏ Р С—Р С•РЎРѓРЎвЂљРЎР‚Р С•Р ВµР Р…Р С‘РЎРЏ [Enter = Р Р†РЎРѓР Вµ]: ");
        string? shield = shieldResult.Status == PromptStatus.OK ? shieldResult.StringResult?.Trim() : null;
        IReadOnlyList<ExcelOutputRow> sourceRows = string.IsNullOrWhiteSpace(shield)
            ? rows
            : rows.Where(x => string.Equals(x.Shield, shield, StringComparison.OrdinalIgnoreCase)).ToList();

        if (sourceRows.Count == 0)
        {
            _log.Write(string.IsNullOrWhiteSpace(shield)
                ? "EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: Р Р…Р ВµРЎвЂљ РЎРѓРЎвЂљРЎР‚Р С•Р С” Р Т‘Р В»РЎРЏ Р С—Р С•РЎРѓРЎвЂљРЎР‚Р С•Р ВµР Р…Р С‘РЎРЏ."
                : $"EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: Р Т‘Р В»РЎРЏ РЎвЂ°Р С‘РЎвЂљР В° '{shield}' РЎРѓРЎвЂљРЎР‚Р С•Р С”Р С‘ Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р…РЎвЂ№.");
            return;
        }

        DrawOlsRows(sourceRows, pointResult.Value);
        _log.Write($"EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: Р С—Р С•РЎРѓРЎвЂљРЎР‚Р С•Р ВµР Р…Р С• РЎРѓРЎвЂљРЎР‚Р С•Р С”: {sourceRows.Count}. Р ВРЎРѓРЎвЂљР С•РЎвЂЎР Р…Р С‘Р С”={(cacheHit ? "Р С”РЎРЊРЎв‚¬" : "РЎвЂћР В°Р в„–Р В»")}.");

        if (workbookExists)
        {
            var tablePrompt = new PromptKeywordOptions("\nР вЂ™РЎРѓРЎвЂљР В°Р Р†Р С‘РЎвЂљРЎРЉ РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…РЎС“РЎР‹ РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎС“ Excel [Р вЂќР В°/Р СњР ВµРЎвЂљ] <Р вЂќР В°>: ");
            tablePrompt.Keywords.Add("Р вЂќР В°");
            tablePrompt.Keywords.Add("Р СњР ВµРЎвЂљ");
            tablePrompt.AllowNone = true;
            PromptResult tableDecision = doc.Editor.GetKeywords(tablePrompt);
            bool shouldInsertTable = tableDecision.Status switch
            {
                PromptStatus.OK => !string.Equals(tableDecision.StringResult, "Р СњР ВµРЎвЂљ", StringComparison.OrdinalIgnoreCase),
                PromptStatus.None => true,
                _ => false
            };

            if (shouldInsertTable)
            {
                PromptPointResult tablePoint = doc.Editor.GetPoint("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ РЎвЂљР С•РЎвЂЎР С”РЎС“ Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р С‘ РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…Р С•Р в„– РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎвЂ№ Excel: ");
                if (tablePoint.Status == PromptStatus.OK)
                {
                    bool tableInserted = _export.TryInsertExcelLinkedTable(templatePath, tablePoint.Value, out string tableStatus);
                    _log.Write($"EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: {tableStatus}");
                    if (!tableInserted)
                    {
                        _log.Write("EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: РЎРѓРЎвЂ¦Р ВµР СР В° Р С—Р С•РЎРѓРЎвЂљРЎР‚Р С•Р ВµР Р…Р В°, Р Р…Р С• Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р В° РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…Р С•Р в„– РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎвЂ№ Р Р…Р Вµ Р Р†РЎвЂ№Р С—Р С•Р В»Р Р…Р ВµР Р…Р В°.");
                    }
                }
                else
                {
                    _log.Write("EOM_Р СџР С›Р РЋР СћР В Р С›Р ВР СћР В¬_Р С›Р вЂєР РЋ: Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р В° РЎРѓР Р†РЎРЏР В·Р В°Р Р…Р Р…Р С•Р в„– РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ РЎвЂ№ Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В° Р С—Р С•Р В»РЎРЉР В·Р С•Р Р†Р В°РЎвЂљР ВµР В»Р ВµР С.");
                }
            }
        }
        // END_BLOCK_COMMAND_EOM_BUILD_OLS
    }

    [CommandMethod(PluginConfig.Commands.BuildPanelLayout)]
    // START_CONTRACT: EomBuildPanelLayout
    //   PURPOSE: Eom build panel layout.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomBuildPanelLayout

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
            MessageForAdding = "\nР вЂ™РЎвЂ№Р Т‘Р ВµР В»Р С‘РЎвЂљР Вµ Р С–Р С•РЎвЂљР С•Р Р†РЎС“РЎР‹ Р С•Р Т‘Р Р…Р С•Р В»Р С‘Р Р…Р ВµР в„–Р Р…РЎС“РЎР‹ РЎРѓРЎвЂ¦Р ВµР СРЎС“ (Р В±Р В»Р С•Р С”Р С‘ Р В°Р С—Р С—Р В°РЎР‚Р В°РЎвЂљР С•Р Р†): "
        };
        SelectionFilter filter = new([
            new TypedValue((int)DxfCode.Start, "INSERT")
        ]);
        PromptSelectionResult selection = doc.Editor.GetSelection(selectionOptions, filter);
        if (selection.Status != PromptStatus.OK)
        {
            _log.Write("EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’ Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°: Р С›Р вЂєР РЋ Р Р…Р Вµ Р Р†РЎвЂ№Р Т‘Р ВµР В»Р ВµР Р…Р В°.");
            return;
        }

        ObjectId[] selectedIds = selection.Value.GetObjectIds();
        if (selectedIds.Length == 0)
        {
            _log.Write("EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’: Р Р…Р Вµ Р Р†РЎвЂ№Р В±РЎР‚Р В°Р Р…РЎвЂ№ Р В±Р В»Р С•Р С”Р С‘ Р С›Р вЂєР РЋ.");
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
            _log.Write("EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’: Р Р…Р ВµРЎвЂљ Р Р†Р В°Р В»Р С‘Р Т‘Р Р…РЎвЂ№РЎвЂ¦ Р С—РЎР‚Р В°Р Р†Р С‘Р В» РЎРѓР С•Р С—Р С•РЎРѓРЎвЂљР В°Р Р†Р В»Р ВµР Р…Р С‘РЎРЏ (Р Р…РЎС“Р В¶Р Р…РЎвЂ№ SelectorRules Р В»Р С‘Р В±Р С• legacy LayoutMap + Р СљР С›Р вЂќР Р€Р вЂєР вЂўР в„ў/FallbackModules).");
            return;
        }

        int defaultModulesPerRow = mapConfig.DefaultModulesPerRow > 0
            ? mapConfig.DefaultModulesPerRow
            : (settings.PanelModulesPerRow > 0 ? settings.PanelModulesPerRow : 24);
        PromptIntegerOptions modulesOptions = new($"\nР СљР С•Р Т‘РЎС“Р В»Р ВµР в„– Р Р† РЎР‚РЎРЏР Т‘РЎС“ [Enter = {defaultModulesPerRow}]: ")
        {
            AllowNone = true,
            LowerLimit = 1,
            UpperLimit = 72
        };
        PromptIntegerResult modulesResult = doc.Editor.GetInteger(modulesOptions);
        if (modulesResult.Status == PromptStatus.Cancel)
        {
            _log.Write("EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’ Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°.");
            return;
        }

        int modulesPerRow = modulesResult.Status == PromptStatus.OK ? modulesResult.Value : defaultModulesPerRow;
        PromptPointResult pointResult = doc.Editor.GetPoint("\nР Р€Р С”Р В°Р В¶Р С‘РЎвЂљР Вµ РЎвЂљР С•РЎвЂЎР С”РЎС“ Р Р†РЎРѓРЎвЂљР В°Р Р†Р С”Р С‘ Р С”Р С•Р СР С—Р С•Р Р…Р С•Р Р†Р С”Р С‘ РЎвЂ°Р С‘РЎвЂљР В°: ");
        if (pointResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’ Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°.");
            return;
        }

        IReadOnlyList<PanelLayoutRow> layoutRows = BuildPanelLayoutModel(mappedDevices, modulesPerRow);
        if (layoutRows.Count == 0)
        {
            _log.Write("EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’: Р Р…Р Вµ РЎРѓРЎвЂћР С•РЎР‚Р СР С‘РЎР‚Р С•Р Р†Р В°Р Р…Р В° Р СР С•Р Т‘Р ВµР В»РЎРЉ РЎР‚Р В°РЎРѓР С”Р В»Р В°Р Т‘Р С”Р С‘.");
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
            $"EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’: Р Р†РЎвЂ№Р Т‘Р ВµР В»Р ВµР Р…Р С• Р В±Р В»Р С•Р С”Р С•Р Р† {selectedIds.Length}, Р Р†Р В°Р В»Р С‘Р Т‘Р Р…РЎвЂ№РЎвЂ¦ Р В±Р В»Р С•Р С”Р С•Р Р† {parsedSelection.Devices.Count}, Р С•РЎвЂљРЎР‚Р С‘РЎРѓР С•Р Р†Р В°Р Р…Р С• РЎС“РЎРѓРЎвЂљРЎР‚Р С•Р в„–РЎРѓРЎвЂљР Р† {uniqueDevices}, РЎРѓР ВµР С–Р СР ВµР Р…РЎвЂљР С•Р Р† {renderedSegments}, РЎР‚РЎРЏР Т‘Р С•Р Р† DIN {dinRows}, Р С—РЎР‚Р С•Р С—РЎС“РЎвЂ°Р ВµР Р…Р С• {issues.Count}.");
        // END_BLOCK_COMMAND_EOM_BUILD_PANEL_LAYOUT
    }

    [CommandMethod(PluginConfig.Commands.BindPanelLayoutVisualization)]
    // START_CONTRACT: EomBindPanelLayoutVisualization
    //   PURPOSE: Eom bind panel layout visualization.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomBindPanelLayoutVisualization

    public void EomBindPanelLayoutVisualization()
    {
        // START_BLOCK_COMMAND_EOM_BIND_PANEL_LAYOUT_VISUALIZATION
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return;
        }

        PanelLayoutMapConfig mapConfig = _settings.LoadPanelLayoutMap();

        var sourceOptions = new PromptEntityOptions("\nР вЂ™РЎвЂ№Р В±Р ВµРЎР‚Р С‘РЎвЂљР Вµ Р С‘РЎРѓРЎвЂ¦Р С•Р Т‘Р Р…РЎвЂ№Р в„– Р В±Р В»Р С•Р С” Р С›Р вЂєР РЋ Р Т‘Р В»РЎРЏ Р С—РЎР‚Р С‘Р Р†РЎРЏР В·Р С”Р С‘: ");
        sourceOptions.SetRejectMessage("\nР СњРЎС“Р В¶Р ВµР Р… Р В±Р В»Р С•Р С” (BlockReference).");
        sourceOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult sourceResult = doc.Editor.GetEntity(sourceOptions);
        if (sourceResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_Р РЋР вЂ™Р Р‡Р вЂ”Р С’Р СћР В¬_Р вЂ™Р ВР вЂ”Р Р€Р С’Р вЂєР ВР вЂ”Р С’Р В¦Р ВР В® Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°: Р С‘РЎРѓРЎвЂ¦Р С•Р Т‘Р Р…РЎвЂ№Р в„– Р В±Р В»Р С•Р С” Р С›Р вЂєР РЋ Р Р…Р Вµ Р Р†РЎвЂ№Р В±РЎР‚Р В°Р Р….");
            return;
        }

        if (!TryReadSourceSignature(doc, sourceResult.ObjectId, out OlsSourceSignature signature))
        {
            _log.Write("EOM_Р РЋР вЂ™Р Р‡Р вЂ”Р С’Р СћР В¬_Р вЂ™Р ВР вЂ”Р Р€Р С’Р вЂєР ВР вЂ”Р С’Р В¦Р ВР В®: Р Р…Р Вµ РЎС“Р Т‘Р В°Р В»Р С•РЎРѓРЎРЉ Р С—РЎР‚Р С•РЎвЂЎР С‘РЎвЂљР В°РЎвЂљРЎРЉ РЎРѓР С‘Р С–Р Р…Р В°РЎвЂљРЎС“РЎР‚РЎС“ Р С‘РЎРѓРЎвЂ¦Р С•Р Т‘Р Р…Р С•Р С–Р С• Р В±Р В»Р С•Р С”Р В°.");
            return;
        }

        var targetOptions = new PromptEntityOptions("\nР вЂ™РЎвЂ№Р В±Р ВµРЎР‚Р С‘РЎвЂљР Вµ Р В±Р В»Р С•Р С” Р Р†Р С‘Р В·РЎС“Р В°Р В»Р С‘Р В·Р В°РЎвЂ Р С‘Р С‘ Р Т‘Р В»РЎРЏ Р С”Р С•Р СР С—Р С•Р Р…Р С•Р Р†Р С”Р С‘ РЎвЂ°Р С‘РЎвЂљР В°: ");
        targetOptions.SetRejectMessage("\nР СњРЎС“Р В¶Р ВµР Р… Р В±Р В»Р С•Р С” (BlockReference).");
        targetOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult targetResult = doc.Editor.GetEntity(targetOptions);
        if (targetResult.Status != PromptStatus.OK)
        {
            _log.Write("EOM_Р РЋР вЂ™Р Р‡Р вЂ”Р С’Р СћР В¬_Р вЂ™Р ВР вЂ”Р Р€Р С’Р вЂєР ВР вЂ”Р С’Р В¦Р ВР В® Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°: Р В±Р В»Р С•Р С” Р Р†Р С‘Р В·РЎС“Р В°Р В»Р С‘Р В·Р В°РЎвЂ Р С‘Р С‘ Р Р…Р Вµ Р Р†РЎвЂ№Р В±РЎР‚Р В°Р Р….");
            return;
        }

        if (!TryReadEffectiveBlockName(doc, targetResult.ObjectId, out string layoutBlockName))
        {
            _log.Write("EOM_Р РЋР вЂ™Р Р‡Р вЂ”Р С’Р СћР В¬_Р вЂ™Р ВР вЂ”Р Р€Р С’Р вЂєР ВР вЂ”Р С’Р В¦Р ВР В®: Р Р…Р Вµ РЎС“Р Т‘Р В°Р В»Р С•РЎРѓРЎРЉ Р С•Р С—РЎР‚Р ВµР Т‘Р ВµР В»Р С‘РЎвЂљРЎРЉ Р С‘Р СРЎРЏ Р В±Р В»Р С•Р С”Р В° Р Р†Р С‘Р В·РЎС“Р В°Р В»Р С‘Р В·Р В°РЎвЂ Р С‘Р С‘.");
            return;
        }

        PromptIntegerOptions priorityOptions = new("\nР СџРЎР‚Р С‘Р С•РЎР‚Р С‘РЎвЂљР ВµРЎвЂљ Р С—РЎР‚Р В°Р Р†Р С‘Р В»Р В° [Enter = 100]: ")
        {
            AllowNone = true,
            LowerLimit = 0,
            UpperLimit = 10000
        };
        PromptIntegerResult priorityResult = doc.Editor.GetInteger(priorityOptions);
        if (priorityResult.Status == PromptStatus.Cancel)
        {
            _log.Write("EOM_Р РЋР вЂ™Р Р‡Р вЂ”Р С’Р СћР В¬_Р вЂ™Р ВР вЂ”Р Р€Р С’Р вЂєР ВР вЂ”Р С’Р В¦Р ВР В® Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°.");
            return;
        }

        int priority = priorityResult.Status == PromptStatus.OK ? priorityResult.Value : 100;
        PromptIntegerOptions fallbackOptions = new("\nFallback Р СР С•Р Т‘РЎС“Р В»Р ВµР в„– [Enter = Р В±Р ВµР В· fallback]: ")
        {
            AllowNone = true,
            LowerLimit = 1,
            UpperLimit = 72
        };
        PromptIntegerResult fallbackResult = doc.Editor.GetInteger(fallbackOptions);
        if (fallbackResult.Status == PromptStatus.Cancel)
        {
            _log.Write("EOM_Р РЋР вЂ™Р Р‡Р вЂ”Р С’Р СћР В¬_Р вЂ™Р ВР вЂ”Р Р€Р С’Р вЂєР ВР вЂ”Р С’Р В¦Р ВР В® Р С•РЎвЂљР СР ВµР Р…Р ВµР Р…Р В°.");
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
            $"EOM_Р РЋР вЂ™Р Р‡Р вЂ”Р С’Р СћР В¬_Р вЂ™Р ВР вЂ”Р Р€Р С’Р вЂєР ВР вЂ”Р С’Р В¦Р ВР В®: РЎРѓР С•РЎвЂ¦РЎР‚Р В°Р Р…Р ВµР Р…Р С• Р С—РЎР‚Р В°Р Р†Р С‘Р В»Р С• SOURCE='{normalizedSourceName}', VISIBILITY='{visibilityPart}' -> LAYOUT='{normalizedLayoutName}', PRIORITY={priority}, FALLBACK_MODULES={fallbackPart}.");
        // END_BLOCK_COMMAND_EOM_BIND_PANEL_LAYOUT_VISUALIZATION
    }

    [CommandMethod(PluginConfig.Commands.BindPanelLayoutVisualizationAlias)]
    // START_CONTRACT: EomBindPanelLayoutVisualizationAlias
    //   PURPOSE: Eom bind panel layout visualization alias.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomBindPanelLayoutVisualizationAlias

    public void EomBindPanelLayoutVisualizationAlias()
    {
        // START_BLOCK_COMMAND_EOM_BIND_PANEL_LAYOUT_VISUALIZATION_ALIAS
        EomBindPanelLayoutVisualization();
        // END_BLOCK_COMMAND_EOM_BIND_PANEL_LAYOUT_VISUALIZATION_ALIAS
    }

    [CommandMethod(PluginConfig.Commands.Validate)]
    // START_CONTRACT: EomValidate
    //   PURPOSE: Eom validate.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EomValidate

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
                        issues.Add(new ValidationIssue("LINE_NO_GROUP", id, "Р вЂєР С‘Р Р…Р С‘РЎРЏ Р В±Р ВµР В· Р вЂњР В Р Р€Р СџР СџР С’", "Error"));
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
                            issues.Add(new ValidationIssue("LINE_DEFAULT_INSTALL_TYPE", id, "Р вЂєР С‘Р Р…Р С‘РЎРЏ РЎРѓ РЎвЂљР С‘Р С—Р С•Р С Р С—РЎР‚Р С•Р С”Р В»Р В°Р Т‘Р С”Р С‘ Р С—Р С• РЎС“Р СР С•Р В»РЎвЂЎР В°Р Р…Р С‘РЎР‹", "Warning"));
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
                        issues.Add(new ValidationIssue("LOAD_NO_GROUP", id, "Р СњР В°Р С–РЎР‚РЎС“Р В·Р С”Р В° Р В±Р ВµР В· Р вЂњР В Р Р€Р СџР СџР С’", "Error"));
                        continue;
                    }

                    string groupValue = group.Trim();
                    loadGroups.Add(groupValue);
                    if (groupRegex is not null && !groupRegex.IsMatch(groupValue))
                    {
                        regexMismatchGroups++;
                        issues.Add(new ValidationIssue("GROUP_REGEX_MISMATCH", id, $"Р СњР ВµРЎРѓРЎвЂљР В°Р Р…Р Т‘Р В°РЎР‚РЎвЂљР Р…РЎвЂ№Р в„– РЎвЂћР С•РЎР‚Р СР В°РЎвЂљ Р вЂњР В Р Р€Р СџР СџР С’: {groupValue}", "Warning"));
                    }

                    if (string.IsNullOrWhiteSpace(power) || !double.TryParse(power.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
                    {
                        invalidPowerLoads++;
                        issues.Add(new ValidationIssue("LOAD_POWER_INVALID", id, "Р СљР С›Р В©Р СњР С›Р РЋР СћР В¬ Р С•РЎвЂљРЎРѓРЎС“РЎвЂљРЎРѓРЎвЂљР Р†РЎС“Р ВµРЎвЂљ Р С‘Р В»Р С‘ Р Р…Р ВµРЎвЂЎР С‘РЎРѓР В»Р С•Р Р†Р В°РЎРЏ", "Error"));
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
                issues.Add(new ValidationIssue("GROUP_WITHOUT_LINES", id, $"Р вЂњРЎР‚РЎС“Р С—Р С—Р В° {group} Р ВµРЎРѓРЎвЂљРЎРЉ РЎС“ Р Р…Р В°Р С–РЎР‚РЎС“Р В·Р С•Р С”, Р Р…Р С• Р Р…Р ВµРЎвЂљ Р В»Р С‘Р Р…Р С‘Р в„–", "Warning"));
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
                issues.Add(new ValidationIssue("GROUP_SHIELD_MISMATCH", id, $"Р вЂњРЎР‚РЎС“Р С—Р С—Р В° {group} РЎРѓР С•Р Т‘Р ВµРЎР‚Р В¶Р С‘РЎвЂљ Р Р…Р ВµРЎРѓР С”Р С•Р В»РЎРЉР С”Р С• РЎвЂ°Р С‘РЎвЂљР С•Р Р†: {string.Join(", ", shields)}", "Warning"));
            }
        }

        _log.Write(
            $"EOM_Р СџР В Р С›Р вЂ™Р вЂўР В Р С™Р С’: Р В»Р С‘Р Р…Р С‘Р С‘ Р В±Р ВµР В· {PluginConfig.Strings.Group}={missingGroupLines}; " +
            $"Р Р…Р В°Р С–РЎР‚РЎС“Р В·Р С”Р С‘ Р В±Р ВµР В· {PluginConfig.Strings.Group}={missingGroupLoads}; " +
            $"{PluginConfig.Strings.Power} Р Р…Р ВµРЎвЂЎР С‘РЎРѓР В»Р С•Р Р†Р В°РЎРЏ={invalidPowerLoads}; " +
            $"Р С–РЎР‚РЎС“Р С—Р С—РЎвЂ№ Р Р…Р В°Р С–РЎР‚РЎС“Р В·Р С•Р С” Р В±Р ВµР В· Р В»Р С‘Р Р…Р С‘Р в„–={groupsWithoutLines}; " +
            $"Р В»Р С‘Р Р…Р С‘Р С‘ РЎРѓ РЎвЂљР С‘Р С—Р С•Р С Р С—Р С• РЎС“Р СР С•Р В»РЎвЂЎР В°Р Р…Р С‘РЎР‹={defaultInstallTypeLines}; " +
            $"Р Р…Р ВµРЎРѓР С•Р Р†Р С—Р В°Р Т‘Р ВµР Р…Р С‘Р Вµ {PluginConfig.Strings.Shield} Р Р†Р Р…РЎС“РЎвЂљРЎР‚Р С‘ Р С–РЎР‚РЎС“Р С—Р С—РЎвЂ№={groupShieldMismatch}; " +
            $"regex-Р С•РЎв‚¬Р С‘Р В±Р С”Р С‘ Р С–РЎР‚РЎС“Р С—Р С—РЎвЂ№={regexMismatchGroups}.");

        if (issues.Count > 0)
        {
            for (int i = 0; i < issues.Count; i++)
            {
                ValidationIssue issue = issues[i];
                _log.Write($"[{i + 1}] {issue.Severity} {issue.Code}: {issue.Message}");
            }

            PromptIntegerOptions pickOptions = new($"\nР вЂ™РЎвЂ№Р В±Р ВµРЎР‚Р С‘РЎвЂљР Вµ Р Р…Р С•Р СР ВµРЎР‚ Р С—РЎР‚Р С•Р В±Р В»Р ВµР СРЎвЂ№ Р Т‘Р В»РЎРЏ Р Р†РЎвЂ№Р Т‘Р ВµР В»Р ВµР Р…Р С‘РЎРЏ [1..{issues.Count}] Р С‘Р В»Р С‘ Enter: ")
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

    // START_CONTRACT: OnObjectAppended
    //   PURPOSE: On object appended.
    //   INPUTS: { objectId: ObjectId - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: OnObjectAppended

    private void OnObjectAppended(ObjectId objectId)
    {
        // START_BLOCK_ON_OBJECT_APPENDED
        _xdata.ApplyActiveGroupToEntity(objectId);
        // END_BLOCK_ON_OBJECT_APPENDED
    }

    // START_CONTRACT: PickPanelLayoutSourceSignatureFromDrawing
    //   PURPOSE: Pick panel layout source signature from drawing.
    //   INPUTS: none
    //   OUTPUTS: { OlsSourceSignature? - result of pick panel layout source signature from drawing }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: PickPanelLayoutSourceSignatureFromDrawing

    private OlsSourceSignature? PickPanelLayoutSourceSignatureFromDrawing()
    {
        // START_BLOCK_PICK_PANEL_LAYOUT_SOURCE_SIGNATURE
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return null;
        }

        var sourceOptions = new PromptEntityOptions("\nР вЂ™РЎвЂ№Р В±Р ВµРЎР‚Р С‘РЎвЂљР Вµ Р С‘РЎРѓРЎвЂ¦Р С•Р Т‘Р Р…РЎвЂ№Р в„– Р В±Р В»Р С•Р С” Р С›Р вЂєР РЋ: ");
        sourceOptions.SetRejectMessage("\nР СњРЎС“Р В¶Р ВµР Р… Р В±Р В»Р С•Р С” (BlockReference).");
        sourceOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult sourceResult = doc.Editor.GetEntity(sourceOptions);
        if (sourceResult.Status != PromptStatus.OK)
        {
            return null;
        }

        if (!TryReadSourceSignature(doc, sourceResult.ObjectId, out OlsSourceSignature signature))
        {
            _log.Write("Р СњР В°РЎРѓРЎвЂљРЎР‚Р С•Р в„–Р С”Р В° Р С”Р С•Р СР С—Р С•Р Р…Р С•Р Р†Р С”Р С‘: Р Р…Р Вµ РЎС“Р Т‘Р В°Р В»Р С•РЎРѓРЎРЉ Р С—РЎР‚Р С•РЎвЂЎР С‘РЎвЂљР В°РЎвЂљРЎРЉ SOURCE (Р С‘Р СРЎРЏ Р В±Р В»Р С•Р С”Р В°/Р Р†Р С‘Р Т‘Р С‘Р СР С•РЎРѓРЎвЂљРЎРЉ).");
            return null;
        }

        return signature;
        // END_BLOCK_PICK_PANEL_LAYOUT_SOURCE_SIGNATURE
    }

    // START_CONTRACT: PickPanelLayoutVisualizationBlockFromDrawing
    //   PURPOSE: Pick panel layout visualization block from drawing.
    //   INPUTS: none
    //   OUTPUTS: { string? - result of pick panel layout visualization block from drawing }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: PickPanelLayoutVisualizationBlockFromDrawing

    private string? PickPanelLayoutVisualizationBlockFromDrawing()
    {
        // START_BLOCK_PICK_PANEL_LAYOUT_VISUALIZATION_BLOCK
        Document? doc = Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return null;
        }

        var targetOptions = new PromptEntityOptions("\nР вЂ™РЎвЂ№Р В±Р ВµРЎР‚Р С‘РЎвЂљР Вµ Р В±Р В»Р С•Р С” Р Р†Р С‘Р В·РЎС“Р В°Р В»Р С‘Р В·Р В°РЎвЂ Р С‘Р С‘: ");
        targetOptions.SetRejectMessage("\nР СњРЎС“Р В¶Р ВµР Р… Р В±Р В»Р С•Р С” (BlockReference).");
        targetOptions.AddAllowedClass(typeof(BlockReference), true);
        PromptEntityResult targetResult = doc.Editor.GetEntity(targetOptions);
        if (targetResult.Status != PromptStatus.OK)
        {
            return null;
        }

        if (!TryReadEffectiveBlockName(doc, targetResult.ObjectId, out string layoutBlockName))
        {
            _log.Write("Р СњР В°РЎРѓРЎвЂљРЎР‚Р С•Р в„–Р С”Р В° Р С”Р С•Р СР С—Р С•Р Р…Р С•Р Р†Р С”Р С‘: Р Р…Р Вµ РЎС“Р Т‘Р В°Р В»Р С•РЎРѓРЎРЉ Р С•Р С—РЎР‚Р ВµР Т‘Р ВµР В»Р С‘РЎвЂљРЎРЉ Р С‘Р СРЎРЏ Р В±Р В»Р С•Р С”Р В° Р Р†Р С‘Р В·РЎС“Р В°Р В»Р С‘Р В·Р В°РЎвЂ Р С‘Р С‘.");
            return null;
        }

        return layoutBlockName.Trim();
        // END_BLOCK_PICK_PANEL_LAYOUT_VISUALIZATION_BLOCK
    }

    // START_CONTRACT: TryResolveExcelTemplatePath
    //   PURPOSE: Attempt to execute resolve excel template path.
    //   INPUTS: { command: string - method parameter; templatePath: out string - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute resolve excel template path }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryResolveExcelTemplatePath

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
                _log.Write($"{command}: Р Р…Р ВµР С”Р С•РЎР‚РЎР‚Р ВµР С”РЎвЂљР Р…РЎвЂ№Р в„– Р С—РЎС“РЎвЂљРЎРЉ РЎв‚¬Р В°Р В±Р В»Р С•Р Р…Р В° Excel. Р СџРЎР‚Р С•Р Р†Р ВµРЎР‚РЎРЉРЎвЂљР Вµ Р С—Р С•Р В»Р Вµ ExcelTemplatePath Р Р† Settings.json.");
                return false;
            }

            return true;
        }
        catch (System.Exception ex)
        {
            _log.Write($"{command}: Р С•РЎв‚¬Р С‘Р В±Р С”Р В° РЎвЂЎРЎвЂљР ВµР Р…Р С‘РЎРЏ Р С—РЎС“РЎвЂљР С‘ РЎв‚¬Р В°Р В±Р В»Р С•Р Р…Р В° Excel: {ex.Message}");
            return false;
        }
        // END_BLOCK_TRY_RESOLVE_EXCEL_TEMPLATE_PATH
    }

    // START_CONTRACT: GetExpectedExcelInputPath
    //   PURPOSE: Retrieve expected excel input path.
    //   INPUTS: { templatePath: string - method parameter }
    //   OUTPUTS: { string - textual result for retrieve expected excel input path }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: GetExpectedExcelInputPath

    private static string GetExpectedExcelInputPath(string templatePath)
    {
        // START_BLOCK_GET_EXPECTED_EXCEL_INPUT_PATH
        string directory = Path.GetDirectoryName(templatePath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        return Path.Combine(string.IsNullOrWhiteSpace(directory) ? "." : directory, $"{fileName}.INPUT.csv");
        // END_BLOCK_GET_EXPECTED_EXCEL_INPUT_PATH
    }

    // START_CONTRACT: GetExpectedExcelOutputPath
    //   PURPOSE: Retrieve expected excel output path.
    //   INPUTS: { templatePath: string - method parameter }
    //   OUTPUTS: { string - textual result for retrieve expected excel output path }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: GetExpectedExcelOutputPath

    private static string GetExpectedExcelOutputPath(string templatePath)
    {
        // START_BLOCK_GET_EXPECTED_EXCEL_OUTPUT_PATH
        string directory = Path.GetDirectoryName(templatePath) ?? string.Empty;
        string fileName = Path.GetFileNameWithoutExtension(templatePath);
        return Path.Combine(string.IsNullOrWhiteSpace(directory) ? "." : directory, $"{fileName}.OUTPUT.csv");
        // END_BLOCK_GET_EXPECTED_EXCEL_OUTPUT_PATH
    }

    // START_CONTRACT: ToExcelInputRow
    //   PURPOSE: To excel input row.
    //   INPUTS: { row: GroupTraceAggregate - method parameter }
    //   OUTPUTS: { ExcelInputRow - result of to excel input row }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ToExcelInputRow

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

    // START_CONTRACT: GetInstallTypeLength
    //   PURPOSE: Retrieve install type length.
    //   INPUTS: { aggregate: GroupTraceAggregate - method parameter; installType: string - method parameter }
    //   OUTPUTS: { double - result of retrieve install type length }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: GetInstallTypeLength

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

    // START_CONTRACT: ReadLineGroup
    //   PURPOSE: Read line group.
    //   INPUTS: { entity: Entity - method parameter }
    //   OUTPUTS: { string? - result of read line group }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ReadLineGroup

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

    // START_CONTRACT: DrawOlsRows
    //   PURPOSE: Draw ols rows.
    //   INPUTS: { rows: IReadOnlyList<ExcelOutputRow> - method parameter; basePoint: Point3d - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: DrawOlsRows

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
                string breakerLabel = string.IsNullOrWhiteSpace(row.CircuitBreaker) ? "QF" : row.CircuitBreaker;
                string rcdLabel = string.IsNullOrWhiteSpace(row.RcdDiff) ? "Р Р€Р вЂ”Р С›" : row.RcdDiff;

                InsertTemplateOrFallback(tr, bt, ms, PluginConfig.TemplateBlocks.Input, inputPoint, "\u0412\u0412\u041e\u0414");
                InsertTemplateOrFallback(tr, bt, ms, PluginConfig.TemplateBlocks.Breaker, breakerPoint, breakerLabel);
                InsertTemplateOrFallback(tr, bt, ms, PluginConfig.TemplateBlocks.Rcd, rcdPoint, rcdLabel);

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

    // START_CONTRACT: WriteOlsRowMetadata
    //   PURPOSE: Write ols row metadata.
    //   INPUTS: { entity: Entity - method parameter; row: ExcelOutputRow - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: WriteOlsRowMetadata

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

    // START_CONTRACT: ReadOlsRowsFromDrawing
    //   PURPOSE: Read ols rows from drawing.
    //   INPUTS: { doc: Document - method parameter }
    //   OUTPUTS: { IReadOnlyList<ExcelOutputRow> - result of read ols rows from drawing }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ReadOlsRowsFromDrawing

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

    // START_CONTRACT: TryParseOlsRowFromEntity
    //   PURPOSE: Attempt to execute parse ols row from entity.
    //   INPUTS: { entity: Entity - method parameter }
    //   OUTPUTS: { ExcelOutputRow? - result of attempt to execute parse ols row from entity }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryParseOlsRowFromEntity

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

    // START_CONTRACT: ToIntOrDefault
    //   PURPOSE: To int or default.
    //   INPUTS: { value: object? - method parameter }
    //   OUTPUTS: { int - result of to int or default }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ToIntOrDefault

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

    // START_CONTRACT: ParseOlsSelection
    //   PURPOSE: Parse ols selection.
    //   INPUTS: { doc: Document - method parameter; selectedIds: IReadOnlyList<ObjectId> - method parameter; tags: PanelLayoutAttributeTags - method parameter }
    //   OUTPUTS: { ParsedOlsSelection - result of parse ols selection }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ParseOlsSelection

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
                    issues.Add(new SkippedOlsDeviceIssue(id, "Р С›Р В±РЎР‰Р ВµР С”РЎвЂљ Р Р…Р Вµ РЎРЏР Р†Р В»РЎРЏР ВµРЎвЂљРЎРѓРЎРЏ BlockReference."));
                    continue;
                }

                IReadOnlyDictionary<string, string> attrs = ReadBlockAttributes(tr, block);
                attrs.TryGetValue(tags.Device, out string? rawDevice);
                attrs.TryGetValue(tags.Modules, out string? rawModules);
                int modules = TryParsePositiveInt(rawModules, out int parsedModules) ? parsedModules : 0;

                string sourceBlockName = ResolveEffectiveBlockName(tr, block);
                if (string.IsNullOrWhiteSpace(sourceBlockName))
                {
                    issues.Add(new SkippedOlsDeviceIssue(id, "Р СњР Вµ РЎС“Р Т‘Р В°Р В»Р С•РЎРѓРЎРЉ Р С•Р С—РЎР‚Р ВµР Т‘Р ВµР В»Р С‘РЎвЂљРЎРЉ Р С‘Р СРЎРЏ Р С‘РЎРѓРЎвЂ¦Р С•Р Т‘Р Р…Р С•Р С–Р С• Р В±Р В»Р С•Р С”Р В° Р С›Р вЂєР РЋ."));
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

    // START_CONTRACT: MapOlsDevices
    //   PURPOSE: Map ols devices.
    //   INPUTS: { devices: IReadOnlyList<OlsSelectedDevice> - method parameter; selectorRules: IReadOnlyList<PanelLayoutSelectorRule> - method parameter; legacyRules: IReadOnlyList<PanelLayoutMapRule> - method parameter; issues: out List<SkippedOlsDeviceIssue> - method parameter }
    //   OUTPUTS: { IReadOnlyList<MappedLayoutDevice> - result of map ols devices }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: MapOlsDevices

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
                    _log.Write($"EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’: SOURCE='{FormatSourceSignature(device.SourceSignature)}' РЎРѓР С•Р Р†Р С—Р В°Р Т‘Р В°Р ВµРЎвЂљ РЎРѓ Р Р…Р ВµРЎРѓР С”Р С•Р В»РЎРЉР С”Р С‘Р СР С‘ Р С—РЎР‚Р В°Р Р†Р С‘Р В»Р В°Р СР С‘ Priority={selectorRule.Priority}. Р СџРЎР‚Р С‘Р СР ВµР Р…Р ВµР Р…Р С• Р С—Р ВµРЎР‚Р Р†Р С•Р Вµ Р С—РЎР‚Р В°Р Р†Р С‘Р В»Р С•.");
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
                    $"Р СњР ВµРЎвЂљ Р С—РЎР‚Р В°Р Р†Р С‘Р В»Р В° PanelLayoutMap.json Р Т‘Р В»РЎРЏ SOURCE='{FormatSourceSignature(device.SourceSignature)}' (Р С‘ fallback Р С—Р С• Р С’Р СџР СџР С’Р В Р С’Р Сћ Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р…).",
                    device.DeviceKey,
                    device.SourceBlockName));
                continue;
            }

            int resolvedModules = ResolveModuleCount(device.Modules, fallbackModules);
            if (resolvedModules <= 0)
            {
                issues.Add(new SkippedOlsDeviceIssue(
                    device.EntityId,
                    $"Р С›РЎвЂљРЎРѓРЎС“РЎвЂљРЎРѓРЎвЂљР Р†РЎС“Р ВµРЎвЂљ Р С”Р С•РЎР‚РЎР‚Р ВµР С”РЎвЂљР Р…РЎвЂ№Р в„– Р В°РЎвЂљРЎР‚Р С‘Р В±РЎС“РЎвЂљ {PluginConfig.PanelLayout.ModulesTag} Р С‘ Р Р† Р С—РЎР‚Р В°Р Р†Р С‘Р В»Р Вµ Р Р…Р ВµРЎвЂљ FallbackModules.",
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

    // START_CONTRACT: ResolveModuleCount
    //   PURPOSE: Resolve module count.
    //   INPUTS: { modulesFromAttribute: int - method parameter; fallbackModules: int? - method parameter }
    //   OUTPUTS: { int - result of resolve module count }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ResolveModuleCount

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

    // START_CONTRACT: ResolveDeviceLabel
    //   PURPOSE: Resolve device label.
    //   INPUTS: { device: OlsSelectedDevice - method parameter }
    //   OUTPUTS: { string - textual result for resolve device label }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ResolveDeviceLabel

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

    // START_CONTRACT: TryResolveSelectorRule
    //   PURPOSE: Attempt to execute resolve selector rule.
    //   INPUTS: { signature: OlsSourceSignature - method parameter; rules: IReadOnlyList<PanelLayoutSelectorRule> - method parameter; samePriorityMatches: out int - method parameter }
    //   OUTPUTS: { PanelLayoutSelectorRule? - result of attempt to execute resolve selector rule }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryResolveSelectorRule

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

    // START_CONTRACT: GetSelectorRuleSpecificity
    //   PURPOSE: Retrieve selector rule specificity.
    //   INPUTS: { rule: PanelLayoutSelectorRule - method parameter }
    //   OUTPUTS: { int - result of retrieve selector rule specificity }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: GetSelectorRuleSpecificity

    private static int GetSelectorRuleSpecificity(PanelLayoutSelectorRule rule)
    {
        // START_BLOCK_GET_SELECTOR_RULE_SPECIFICITY
        return string.IsNullOrWhiteSpace(rule.VisibilityValue) ? 0 : 1;
        // END_BLOCK_GET_SELECTOR_RULE_SPECIFICITY
    }

    // START_CONTRACT: IsSelectorRuleMatch
    //   PURPOSE: Check whether selector rule match.
    //   INPUTS: { rule: PanelLayoutSelectorRule - method parameter; signature: OlsSourceSignature - method parameter }
    //   OUTPUTS: { bool - true when method can check whether selector rule match }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: IsSelectorRuleMatch

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

    // START_CONTRACT: FormatSourceSignature
    //   PURPOSE: Format source signature.
    //   INPUTS: { signature: OlsSourceSignature - method parameter }
    //   OUTPUTS: { string - textual result for format source signature }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: FormatSourceSignature

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
    //   PURPOSE: Build split-aware DIN row placement model for EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’ from mapped OLS devices.
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

    // START_CONTRACT: NormalizeModuleCount
    //   PURPOSE: Normalize module count.
    //   INPUTS: { rawModuleCount: int - method parameter }
    //   OUTPUTS: { int - result of normalize module count }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: NormalizeModuleCount

    private static int NormalizeModuleCount(int rawModuleCount)
    {
        // START_BLOCK_NORMALIZE_MODULE_COUNT
        return rawModuleCount <= 0 ? 1 : rawModuleCount;
        // END_BLOCK_NORMALIZE_MODULE_COUNT
    }

    // START_CONTRACT: CountSegmentsForDevice
    //   PURPOSE: Count segments for device.
    //   INPUTS: { totalModules: int - method parameter; occupiedInRow: int - method parameter; modulesPerRow: int - method parameter }
    //   OUTPUTS: { int - result of count segments for device }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: CountSegmentsForDevice

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

    // START_CONTRACT: DrawPanelLayout
    //   PURPOSE: Draw panel layout.
    //   INPUTS: { layoutRows: IReadOnlyList<PanelLayoutRow> - method parameter; basePoint: Point3d - method parameter; modulesPerRow: int - method parameter; issues: out List<SkippedOlsDeviceIssue> - method parameter }
    //   OUTPUTS: { int - result of draw panel layout }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: DrawPanelLayout

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
                        $"Р вЂ™ РЎвЂЎР ВµРЎР‚РЎвЂљР ВµР В¶Р Вµ Р С•РЎвЂљРЎРѓРЎС“РЎвЂљРЎРѓРЎвЂљР Р†РЎС“Р ВµРЎвЂљ Р В±Р В»Р С•Р С” Р Р†Р С‘Р В·РЎС“Р В°Р В»Р С‘Р В·Р В°РЎвЂ Р С‘Р С‘ '{row.LayoutBlockName}'.",
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

    // START_CONTRACT: DrawPanelLayoutRowFrame
    //   PURPOSE: Draw panel layout row frame.
    //   INPUTS: { ms: BlockTableRecord - method parameter; tr: Transaction - method parameter; rowStartX: double - method parameter; rowTopY: double - method parameter; z: double - method parameter; modulesPerRow: int - method parameter; moduleWidth: double - method parameter; moduleHeight: double - method parameter; dinRow: int - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: DrawPanelLayoutRowFrame

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
            TextString = $"Р В РЎРЏР Т‘ {dinRow}",
            Layer = "0"
        };
        ms.AppendEntity(rowLabel);
        tr.AddNewlyCreatedDBObject(rowLabel, true);
        // END_BLOCK_DRAW_PANEL_LAYOUT_ROW_FRAME
    }

    // START_CONTRACT: DrawPanelLayoutDevice
    //   PURPOSE: Draw panel layout device.
    //   INPUTS: { ms: BlockTableRecord - method parameter; tr: Transaction - method parameter; bt: BlockTable - method parameter; row: PanelLayoutRow - method parameter; leftX: double - method parameter; topY: double - method parameter; z: double - method parameter; moduleWidth: double - method parameter; moduleHeight: double - method parameter }
    //   OUTPUTS: { bool - true when method can draw panel layout device }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: DrawPanelLayoutDevice

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

    // START_CONTRACT: TryInsertLayoutBlockFitted
    //   PURPOSE: Attempt to execute insert layout block fitted.
    //   INPUTS: { ms: BlockTableRecord - method parameter; tr: Transaction - method parameter; bt: BlockTable - method parameter; blockName: string - method parameter; leftX: double - method parameter; topY: double - method parameter; z: double - method parameter; targetWidth: double - method parameter; targetHeight: double - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute insert layout block fitted }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryInsertLayoutBlockFitted

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

    // START_CONTRACT: TryGetBlockDefinitionExtents
    //   PURPOSE: Attempt to execute get block definition extents.
    //   INPUTS: { tr: Transaction - method parameter; btr: BlockTableRecord - method parameter; minPoint: out Point3d - method parameter; maxPoint: out Point3d - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute get block definition extents }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryGetBlockDefinitionExtents

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

    // START_CONTRACT: DrawRectangle
    //   PURPOSE: Draw rectangle.
    //   INPUTS: { ms: BlockTableRecord - method parameter; tr: Transaction - method parameter; leftX: double - method parameter; topY: double - method parameter; width: double - method parameter; height: double - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: DrawRectangle

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

    // START_CONTRACT: ResolveBlockName
    //   PURPOSE: Resolve block name.
    //   INPUTS: { tr: Transaction - method parameter; block: BlockReference - method parameter }
    //   OUTPUTS: { string - textual result for resolve block name }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ResolveBlockName

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

    // START_CONTRACT: ResolveEffectiveBlockName
    //   PURPOSE: Resolve effective block name.
    //   INPUTS: { tr: Transaction - method parameter; block: BlockReference - method parameter }
    //   OUTPUTS: { string - textual result for resolve effective block name }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ResolveEffectiveBlockName

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

    // START_CONTRACT: TryResolveOwnerDynamicBlockName
    //   PURPOSE: Attempt to execute resolve owner dynamic block name.
    //   INPUTS: { tr: Transaction - method parameter; block: BlockReference - method parameter }
    //   OUTPUTS: { string? - result of attempt to execute resolve owner dynamic block name }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryResolveOwnerDynamicBlockName

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

    // START_CONTRACT: IsAnonymousBlockName
    //   PURPOSE: Check whether anonymous block name.
    //   INPUTS: { name: string? - method parameter }
    //   OUTPUTS: { bool - true when method can check whether anonymous block name }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: IsAnonymousBlockName

    private static bool IsAnonymousBlockName(string? name)
    {
        // START_BLOCK_IS_ANONYMOUS_BLOCK_NAME
        return !string.IsNullOrWhiteSpace(name) && name.StartsWith("*", StringComparison.Ordinal);
        // END_BLOCK_IS_ANONYMOUS_BLOCK_NAME
    }

    // START_CONTRACT: ResolveVisibilityValue
    //   PURPOSE: Resolve visibility value.
    //   INPUTS: { block: BlockReference - method parameter }
    //   OUTPUTS: { string? - result of resolve visibility value }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ResolveVisibilityValue

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

    // START_CONTRACT: IsVisibilityLikeProperty
    //   PURPOSE: Check whether visibility like property.
    //   INPUTS: { propertyName: string? - method parameter }
    //   OUTPUTS: { bool - true when method can check whether visibility like property }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: IsVisibilityLikeProperty

    private static bool IsVisibilityLikeProperty(string? propertyName)
    {
        // START_BLOCK_IS_VISIBILITY_LIKE_PROPERTY
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            return false;
        }

        return propertyName.Contains("visibility", StringComparison.OrdinalIgnoreCase)
            || propertyName.Contains("Р Р†Р С‘Р Т‘Р С‘Р С", StringComparison.OrdinalIgnoreCase)
            || propertyName.Contains("lookup", StringComparison.OrdinalIgnoreCase)
            || propertyName.Contains("РЎвЂљР В°Р В±Р В»Р С‘РЎвЂ ", StringComparison.OrdinalIgnoreCase);
        // END_BLOCK_IS_VISIBILITY_LIKE_PROPERTY
    }

    // START_CONTRACT: TryReadSourceSignature
    //   PURPOSE: Attempt to execute read source signature.
    //   INPUTS: { doc: Document - method parameter; objectId: ObjectId - method parameter; signature: out OlsSourceSignature - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute read source signature }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryReadSourceSignature

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

    // START_CONTRACT: TryReadEffectiveBlockName
    //   PURPOSE: Attempt to execute read effective block name.
    //   INPUTS: { doc: Document - method parameter; objectId: ObjectId - method parameter; blockName: out string - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute read effective block name }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryReadEffectiveBlockName

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

    // START_CONTRACT: TryParsePositiveInt
    //   PURPOSE: Attempt to execute parse positive int.
    //   INPUTS: { raw: string? - method parameter; value: out int - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute parse positive int }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryParsePositiveInt

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

    // START_CONTRACT: ReportSkippedOlsDevices
    //   PURPOSE: Report skipped ols devices.
    //   INPUTS: { issues: IReadOnlyList<SkippedOlsDeviceIssue> - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: ReportSkippedOlsDevices

    private void ReportSkippedOlsDevices(IReadOnlyList<SkippedOlsDeviceIssue> issues)
    {
        // START_BLOCK_REPORT_SKIPPED_OLS_DEVICES
        if (issues.Count == 0)
        {
            return;
        }

        _log.Write($"EOM_Р С™Р С›Р СљР СџР С›Р СњР С›Р вЂ™Р С™Р С’_Р В©Р ВР СћР С’: Р С—РЎР‚Р С•Р С—РЎС“РЎвЂ°Р ВµР Р…Р С• Р С•Р В±РЎР‰Р ВµР С”РЎвЂљР С•Р Р†: {issues.Count}.");
        foreach (SkippedOlsDeviceIssue issue in issues)
        {
            string blockNamePart = string.IsNullOrWhiteSpace(issue.SourceBlockName) ? string.Empty : $" BLOCK={issue.SourceBlockName};";
            string deviceKeyPart = string.IsNullOrWhiteSpace(issue.DeviceKey) ? string.Empty : $" Р С’Р СџР СџР С’Р В Р С’Р Сћ={issue.DeviceKey};";
            _log.Write($"  - ID={issue.EntityId};{blockNamePart}{deviceKeyPart} Р СџРЎР‚Р С‘РЎвЂЎР С‘Р Р…Р В°: {issue.Reason}");
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

    // START_CONTRACT: InsertTemplateOrFallback
    //   PURPOSE: Insert template or fallback.
    //   INPUTS: { tr: Transaction - method parameter; bt: BlockTable - method parameter; ms: BlockTableRecord - method parameter; blockName: string - method parameter; position: Point3d - method parameter; fallbackLabel: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: InsertTemplateOrFallback

    private static void InsertTemplateOrFallback(Transaction tr, BlockTable bt, BlockTableRecord ms, string blockName, Point3d position, string fallbackLabel)
    {
        // START_BLOCK_INSERT_TEMPLATE_OR_FALLBACK
        if (TryResolveTemplateBlockId(tr, bt, blockName, out ObjectId blockId))
        {
            var blockRef = new BlockReference(position, blockId);
            ms.AppendEntity(blockRef);
            tr.AddNewlyCreatedDBObject(blockRef, true);
            return;
        }

        var fallback = new DBText
        {
            Position = position,
            Height = 2.5,
            TextString = string.IsNullOrWhiteSpace(fallbackLabel) ? blockName : fallbackLabel,
            Layer = "0"
        };
        ms.AppendEntity(fallback);
        tr.AddNewlyCreatedDBObject(fallback, true);
        // END_BLOCK_INSERT_TEMPLATE_OR_FALLBACK
    }

    // START_CONTRACT: TryResolveTemplateBlockId
    //   PURPOSE: Attempt to execute resolve template block id.
    //   INPUTS: { tr: Transaction - method parameter; bt: BlockTable - method parameter; requestedName: string - method parameter; blockId: out ObjectId - method parameter }
    //   OUTPUTS: { bool - true when method can attempt to execute resolve template block id }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: TryResolveTemplateBlockId

    private static bool TryResolveTemplateBlockId(Transaction tr, BlockTable bt, string requestedName, out ObjectId blockId)
    {
        // START_BLOCK_TRY_RESOLVE_TEMPLATE_BLOCK_ID
        if (!string.IsNullOrWhiteSpace(requestedName) && bt.Has(requestedName))
        {
            blockId = bt[requestedName];
            return true;
        }

        string requestedKey = NormalizeBlockNameForLookup(requestedName);
        if (string.IsNullOrWhiteSpace(requestedKey))
        {
            blockId = ObjectId.Null;
            return false;
        }

        IReadOnlyList<string> tokens = GetBlockSearchTokens(requestedName);
        ObjectId exactNormalizedMatch = ObjectId.Null;
        ObjectId bestTokenMatch = ObjectId.Null;
        int bestTokenScore = int.MaxValue;

        foreach (ObjectId candidateId in bt)
        {
            if (tr.GetObject(candidateId, OpenMode.ForRead) is not BlockTableRecord candidate)
            {
                continue;
            }

            if (candidate.IsLayout || candidate.IsAnonymous || candidate.IsDependent)
            {
                continue;
            }

            string candidateName = candidate.Name ?? string.Empty;
            if (string.IsNullOrWhiteSpace(candidateName))
            {
                continue;
            }

            string candidateKey = NormalizeBlockNameForLookup(candidateName);
            if (string.Equals(candidateKey, requestedKey, StringComparison.OrdinalIgnoreCase))
            {
                exactNormalizedMatch = candidateId;
                break;
            }

            int tokenHits = tokens.Count(token => candidateKey.Contains(token, StringComparison.OrdinalIgnoreCase));
            if (tokenHits <= 0)
            {
                continue;
            }

            int score = candidateKey.Length - (tokenHits * 100);
            if (candidateKey.Contains("Р С•Р В»РЎРѓ", StringComparison.OrdinalIgnoreCase))
            {
                score -= 50;
            }

            if (score < bestTokenScore)
            {
                bestTokenScore = score;
                bestTokenMatch = candidateId;
            }
        }

        if (!exactNormalizedMatch.IsNull)
        {
            blockId = exactNormalizedMatch;
            return true;
        }

        if (!bestTokenMatch.IsNull)
        {
            blockId = bestTokenMatch;
            return true;
        }

        blockId = ObjectId.Null;
        return false;
        // END_BLOCK_TRY_RESOLVE_TEMPLATE_BLOCK_ID
    }

    // START_CONTRACT: GetBlockSearchTokens
    //   PURPOSE: Retrieve block search tokens.
    //   INPUTS: { blockName: string - method parameter }
    //   OUTPUTS: { IReadOnlyList<string> - result of retrieve block search tokens }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: GetBlockSearchTokens

    private static IReadOnlyList<string> GetBlockSearchTokens(string blockName)
    {
        // START_BLOCK_GET_BLOCK_SEARCH_TOKENS
        string normalized = NormalizeBlockNameForLookup(blockName);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return [];
        }

        string breakerKey = NormalizeBlockNameForLookup(PluginConfig.TemplateBlocks.Breaker);
        if (string.Equals(normalized, breakerKey, StringComparison.OrdinalIgnoreCase))
        {
            return ["Р В°Р Р†РЎвЂљР С•Р СР В°РЎвЂљ", "qf", "Р Р†Р В°"];
        }

        string rcdKey = NormalizeBlockNameForLookup(PluginConfig.TemplateBlocks.Rcd);
        if (string.Equals(normalized, rcdKey, StringComparison.OrdinalIgnoreCase))
        {
            return ["РЎС“Р В·Р С•", "Р Т‘Р С‘РЎвЂћ", "Р В°Р Р†Р Т‘РЎвЂљ", "rcd"];
        }

        string inputKey = NormalizeBlockNameForLookup(PluginConfig.TemplateBlocks.Input);
        if (string.Equals(normalized, inputKey, StringComparison.OrdinalIgnoreCase))
        {
            return ["Р Р†Р Р†Р С•Р Т‘", "qs"];
        }

        return [normalized];
        // END_BLOCK_GET_BLOCK_SEARCH_TOKENS
    }

    // START_CONTRACT: NormalizeBlockNameForLookup
    //   PURPOSE: Normalize block name for lookup.
    //   INPUTS: { value: string - method parameter }
    //   OUTPUTS: { string - textual result for normalize block name for lookup }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: NormalizeBlockNameForLookup

    private static string NormalizeBlockNameForLookup(string value)
    {
        // START_BLOCK_NORMALIZE_BLOCK_NAME_FOR_LOOKUP
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length);
        foreach (char ch in value.Trim().ToLowerInvariant())
        {
            if (char.IsLetterOrDigit(ch))
            {
                builder.Append(ch);
            }
        }

        return builder.ToString();
        // END_BLOCK_NORMALIZE_BLOCK_NAME_FOR_LOOKUP
    }

    // START_CONTRACT: EnsureRegApp
    //   PURPOSE: Ensure reg app.
    //   INPUTS: { tr: Transaction - method parameter; db: Database - method parameter; appName: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: EnsureRegApp

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

    // START_CONTRACT: UpsertEntityXData
    //   PURPOSE: Upsert entity X data.
    //   INPUTS: { entity: Entity - method parameter; appName: string - method parameter; appPayload: IReadOnlyList<TypedValue> - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: UpsertEntityXData

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

    // START_CONTRACT: SelectIssueEntity
    //   PURPOSE: Select issue entity.
    //   INPUTS: { editor: Editor - method parameter; issue: ValidationIssue - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-ENTRY-COMMANDS
    // END_CONTRACT: SelectIssueEntity

    private static void SelectIssueEntity(Editor editor, ValidationIssue issue)
    {
        // START_BLOCK_SELECT_ISSUE_ENTITY
        if (issue.EntityId.IsNull)
        {
            return;
        }

        editor.SetImpliedSelection(new[] { issue.EntityId });
        editor.WriteMessage($"\nР вЂ™РЎвЂ№Р В±РЎР‚Р В°Р Р… Р С•Р В±РЎР‰Р ВµР С”РЎвЂљ Р С—РЎР‚Р С•Р В±Р В»Р ВµР СРЎвЂ№: {issue.Code}");
        // END_BLOCK_SELECT_ISSUE_ENTITY
    }
}