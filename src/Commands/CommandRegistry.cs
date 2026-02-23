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
using ElTools.Integrations;
using ElTools.Models;
using ElTools.Services;
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
}
