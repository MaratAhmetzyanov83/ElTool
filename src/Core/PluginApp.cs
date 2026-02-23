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

    public void Initialize()
    {
        // START_BLOCK_PLUGIN_INITIALIZE
        var ribbon = new RibbonBuilder();
        ribbon.BuildRibbon();
        _log.Write("Плагин ElTools инициализирован.");
        // END_BLOCK_PLUGIN_INITIALIZE
    }

    public void Terminate()
    {
        // START_BLOCK_PLUGIN_TERMINATE
        _log.Write("Плагин ElTools завершает работу.");
        // END_BLOCK_PLUGIN_TERMINATE
    }
}
