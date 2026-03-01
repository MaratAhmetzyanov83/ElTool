// FILE: src/Services/LogService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Write operational diagnostics to AutoCAD editor and optional sinks.
//   SCOPE: Contextual logging for plugin operations.
//   DEPENDS: M-CAD-CONTEXT
//   LINKS: M-LOGGING, M-CAD-CONTEXT
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Write - Writes prefixed message to AutoCAD editor.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using ElTools.Integrations;

namespace ElTools.Services;

public class LogService
{
    private readonly AutoCADAdapter _acad = new();

    // START_CONTRACT: Write
    //   PURPOSE: Write.
    //   INPUTS: { message: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-LOGGING
    // END_CONTRACT: Write

    public void Write(string message)
    {
        // START_BLOCK_WRITE_EDITOR_LOG
        Document? doc = _acad.GetActiveDocument();
        if (doc is null)
        {
            return;
        }

        doc.Editor.WriteMessage($"\n[ElTools] {message}");
        // END_BLOCK_WRITE_EDITOR_LOG
    }
}