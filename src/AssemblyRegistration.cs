// FILE: src/AssemblyRegistration.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Register extension entry point and command container for AutoCAD loader.
//   SCOPE: Assembly-level attributes for plugin bootstrap and command discovery.
//   DEPENDS: M-ENTRY, M-COMMANDS
//   LINKS: M-ENTRY, M-COMMANDS
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Assembly attributes - ExtensionApplication and CommandClass registration.
// END_MODULE_MAP

using Autodesk.AutoCAD.Runtime;
using ElTools.Commands;
using ElTools.Core;

[assembly: ExtensionApplication(typeof(PluginApp))]
[assembly: CommandClass(typeof(CommandRegistry))]
