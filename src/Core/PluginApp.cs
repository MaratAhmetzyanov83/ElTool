// FILE: src/Core/PluginApp.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Bootstrap plugin lifecycle and compose startup services.
//   SCOPE: Initialize ribbon and command entry infrastructure.
//   DEPENDS: M-RIBBON, M-COMMANDS, M-LOG
//   LINKS: M-ENTRY, M-RIBBON, M-COMMANDS, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Initialize - Starts plugin modules.
//   Terminate - Finalizes plugin lifecycle.
// END_MODULE_MAP

using Autodesk.AutoCAD.Runtime;
using ElTools.Services;

namespace ElTools.Core;

public class PluginApp : IExtensionApplication
{
    private readonly LogService _log = new();
    private readonly RibbonBuilder _ribbonBuilder = new();
    private bool _isRibbonInitialized;

    public void Initialize()
    {
        // START_BLOCK_PLUGIN_INITIALIZE
        TryBuildRibbon();
        if (!_isRibbonInitialized)
        {
            Application.Idle += OnApplicationIdle;
        }

        _log.Write("Плагин ElTools инициализирован.");
        // END_BLOCK_PLUGIN_INITIALIZE
    }

    public void Terminate()
    {
        // START_BLOCK_PLUGIN_TERMINATE
        Application.Idle -= OnApplicationIdle;
        _log.Write("Плагин ElTools завершает работу.");
        // END_BLOCK_PLUGIN_TERMINATE
    }

    private void OnApplicationIdle(object? sender, EventArgs e)
    {
        // START_BLOCK_PLUGIN_IDLE_RIBBON_INIT
        TryBuildRibbon();
        // END_BLOCK_PLUGIN_IDLE_RIBBON_INIT
    }

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
            _log.Write("Лента ElTools инициализирована.");
        }
        // END_BLOCK_PLUGIN_TRY_BUILD_RIBBON
    }
}
