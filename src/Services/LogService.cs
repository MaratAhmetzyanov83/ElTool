// FILE: src/Services/LogService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Write operational diagnostics to AutoCAD editor and optional sinks.
//   SCOPE: Contextual logging for plugin operations.
//   DEPENDS: M-ACAD
//   LINKS: M-LOG, M-ACAD
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Write - Writes prefixed message to AutoCAD editor.
// END_MODULE_MAP

using ElTools.Integrations;

namespace ElTools.Services;

public class LogService
{
    private readonly AutoCADAdapter _acad = new();

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
