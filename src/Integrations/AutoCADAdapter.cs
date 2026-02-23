// FILE: src/Integrations/AutoCADAdapter.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Provide safe access to AutoCAD document, editor, database, transactions, and coordinate conversions.
//   SCOPE: Active document access, transaction wrapper, UCS->WCS conversion.
//   DEPENDS: none
//   LINKS: M-ACAD
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   GetActiveDocument - Returns active AutoCAD document safely.
//   RunTransaction - Executes action inside database transaction.
//   UcsToWcs - Converts points from UCS to WCS.
// END_MODULE_MAP

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace ElTools.Integrations;

public class AutoCADAdapter
{
    public Document? GetActiveDocument()
    {
        // START_BLOCK_GET_ACTIVE_DOCUMENT
        return Application.DocumentManager.MdiActiveDocument;
        // END_BLOCK_GET_ACTIVE_DOCUMENT
    }

    public void RunTransaction(Action<Transaction, Database> operation)
    {
        // START_BLOCK_RUN_TRANSACTION
        Document? doc = GetActiveDocument();
        if (doc is null)
        {
            return;
        }

        Database db = doc.Database;
        using Transaction tr = db.TransactionManager.StartTransaction();
        operation(tr, db);
        tr.Commit();
        // END_BLOCK_RUN_TRANSACTION
    }

    public Point3d UcsToWcs(Point3d point)
    {
        // START_BLOCK_UCS_TO_WCS
        Document? doc = GetActiveDocument();
        if (doc is null)
        {
            return point;
        }

        return point.TransformBy(doc.Editor.CurrentUserCoordinateSystem);
        // END_BLOCK_UCS_TO_WCS
    }
}
