// FILE: src/Core/PluginApp.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Bootstrap plugin lifecycle and compose startup services.
//   SCOPE: Initialize ribbon and command entry infrastructure.
//   DEPENDS: M-RIBBON, M-ENTRY-COMMANDS, M-LOGGING
//   LINKS: M-PLUGIN-LIFECYCLE, M-RIBBON, M-ENTRY-COMMANDS, M-LOGGING
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Initialize - Starts plugin modules.
//   Terminate - Finalizes plugin lifecycle.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using Autodesk.AutoCAD.Runtime;
using ElTools.Services;

namespace ElTools.Core;

public class PluginApp : IExtensionApplication
{
    private readonly LogService _log = new();
    private readonly RibbonBuilder _ribbonBuilder = new();
    private bool _isRibbonInitialized;

    // START_CONTRACT: Initialize
    //   PURPOSE: Initialize.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-PLUGIN-LIFECYCLE
    // END_CONTRACT: Initialize

    public void Initialize()
    {
        // START_BLOCK_PLUGIN_INITIALIZE
        TryBuildRibbon();
        if (!_isRibbonInitialized)
        {
            Application.Idle += OnApplicationIdle;
        }

        _log.Write("Р В Р’В Р РЋРЎСџР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂњР В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В¦ ElTools Р В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В°Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В·Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В¦.");
        // END_BLOCK_PLUGIN_INITIALIZE
    }

    // START_CONTRACT: Terminate
    //   PURPOSE: Terminate.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-PLUGIN-LIFECYCLE
    // END_CONTRACT: Terminate

    public void Terminate()
    {
        // START_BLOCK_PLUGIN_TERMINATE
        Application.Idle -= OnApplicationIdle;
        _log.Write("Р В Р’В Р РЋРЎСџР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂњР В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В¦ ElTools Р В Р’В Р вЂ™Р’В·Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РІР‚С™Р В Р Р‹Р Р†РІР‚С™Р’В¬Р В Р’В Р вЂ™Р’В°Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р Р†Р вЂљРЎв„ў Р В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р вЂ™Р’В±Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р РЋРІР‚Сљ.");
        // END_BLOCK_PLUGIN_TERMINATE
    }

    // START_CONTRACT: OnApplicationIdle
    //   PURPOSE: On application idle.
    //   INPUTS: { sender: object? - method parameter; e: EventArgs - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-PLUGIN-LIFECYCLE
    // END_CONTRACT: OnApplicationIdle

    private void OnApplicationIdle(object? sender, EventArgs e)
    {
        // START_BLOCK_PLUGIN_IDLE_RIBBON_INIT
        TryBuildRibbon();
        // END_BLOCK_PLUGIN_IDLE_RIBBON_INIT
    }

    // START_CONTRACT: TryBuildRibbon
    //   PURPOSE: Attempt to execute build ribbon.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-PLUGIN-LIFECYCLE
    // END_CONTRACT: TryBuildRibbon

    private void TryBuildRibbon()
    {
        // START_BLOCK_PLUGIN_TRY_BUILD_RIBBON
        if (_isRibbonInitialized)
        {
            return;
        }

        _isRibbonInitialized = _ribbonBuilder.BuildRibbon();
        if (_isRibbonInitialized)
        {
            Application.Idle -= OnApplicationIdle;
            _log.Write("Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В° ElTools Р В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В°Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В·Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’В°.");
        }
        // END_BLOCK_PLUGIN_TRY_BUILD_RIBBON
    }
}