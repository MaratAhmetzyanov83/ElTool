// FILE: src/AssemblyRegistration.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Register extension entry point and command container for AutoCAD loader.
//   SCOPE: Assembly-level attributes for plugin bootstrap and command discovery.
//   DEPENDS: M-PLUGIN-LIFECYCLE, M-ENTRY-COMMANDS
//   LINKS: M-ASSEMBLY-REGISTRATION, M-PLUGIN-LIFECYCLE, M-ENTRY-COMMANDS
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Assembly attributes - ExtensionApplication and CommandClass registration.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using Autodesk.AutoCAD.Runtime;
using ElTools.Commands;
using ElTools.Core;

[assembly: ExtensionApplication(typeof(PluginApp))]
[assembly: CommandClass(typeof(CommandRegistry))]