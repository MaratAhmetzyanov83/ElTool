// FILE: src/Services/SpecificationService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Scan drawing EOM_DATA and aggregate cable lengths by type and group.
//   SCOPE: Reads XData records and groups by type/group.
//   DEPENDS: M-XDATA, M-SETTINGS, M-LOGGING, M-MODELS
//   LINKS: M-AGGREGATION, M-XDATA, M-SETTINGS, M-LOGGING, M-MODELS
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   BuildSpecification - Produces grouped specification rows.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using ElTools.Data;
using ElTools.Models;

namespace ElTools.Services;

public class SpecificationService
{
    private readonly XDataService _xdata = new();
    private readonly SettingsRepository _settings = new();
    private readonly LogService _log = new();

    // START_CONTRACT: BuildSpecification
    //   PURPOSE: Build specification.
    //   INPUTS: { entityIds: IEnumerable<ObjectId> - method parameter }
    //   OUTPUTS: { IReadOnlyList<SpecificationRow> - result of build specification }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-AGGREGATION
    // END_CONTRACT: BuildSpecification

    public IReadOnlyList<SpecificationRow> BuildSpecification(IEnumerable<ObjectId> entityIds)
    {
        // START_BLOCK_BUILD_SPECIFICATION
        _ = _settings.LoadSettings();

        List<ObjectId> sourceIds = entityIds.ToList();
        if (sourceIds.Count == 0)
        {
            sourceIds = CollectModelSpaceEntityIds();
        }

        var records = new List<EomDataRecord>();
        foreach (ObjectId id in sourceIds)
        {
            EomDataRecord? item = _xdata.ReadEomData(id);
            if (item is not null)
            {
                records.Add(item);
            }
        }

        var grouped = records
            .GroupBy(x => new { x.Type, x.Group })
            .Select(g => new SpecificationRow(g.Key.Type, g.Key.Group, Math.Round(g.Sum(x => x.TotalLength), 3)))
            .OrderBy(x => x.CableType)
            .ThenBy(x => x.Group)
            .ToList();

        _log.Write($"Р РЋР С—Р ВµРЎвЂ Р С‘РЎвЂћР С‘Р С”Р В°РЎвЂ Р С‘РЎРЏ РЎРѓР С•Р В±РЎР‚Р В°Р Р…Р В°. Р РЋРЎвЂљРЎР‚Р С•Р С”: {grouped.Count}.");
        return grouped;
        // END_BLOCK_BUILD_SPECIFICATION
    }

    // START_CONTRACT: CollectModelSpaceEntityIds
    //   PURPOSE: Collect model space entity ids.
    //   INPUTS: none
    //   OUTPUTS: { List<ObjectId> - result of collect model space entity ids }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-AGGREGATION
    // END_CONTRACT: CollectModelSpaceEntityIds

    private static List<ObjectId> CollectModelSpaceEntityIds()
    {
        // START_BLOCK_COLLECT_MODEL_SPACE_ENTITY_IDS
        Document? doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return [];
        }

        var ids = new List<ObjectId>();
        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            var blockTable = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
            ids.AddRange(modelSpace.Cast<ObjectId>());
            tr.Commit();
        }

        return ids;
        // END_BLOCK_COLLECT_MODEL_SPACE_ENTITY_IDS
    }
}