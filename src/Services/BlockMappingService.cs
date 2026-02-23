// FILE: src/Services/BlockMappingService.cs
// VERSION: 1.1.0
// START_MODULE_CONTRACT
//   PURPOSE: Replace source blocks with a selected target block in drawing.
//   SCOPE: Source/target template detection, in-place block replacement, replacement summary.
//   DEPENDS: M-ACAD, M-LOG
//   LINKS: M-MAP, M-ACAD, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   ExecuteMapping - Replaces all source block instances by selected target block definition.
// END_MODULE_MAP
//
// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.1.0 - Switched mapping workflow to direct source/target selection without settings profile and attribute transfer.
// END_CHANGE_SUMMARY

using ElTools.Integrations;

namespace ElTools.Services;

public class BlockMappingService
{
    private readonly AutoCADAdapter _acad = new();
    private readonly LogService _log = new();

    public int ExecuteMapping(ObjectId sourceSampleId, ObjectId targetSampleId)
    {
        // START_BLOCK_EXECUTE_MAPPING
        int replaced = 0;
        _acad.RunTransaction((tr, db) =>
        {
            var sourceSample = tr.GetObject(sourceSampleId, OpenMode.ForRead) as BlockReference;
            var targetSample = tr.GetObject(targetSampleId, OpenMode.ForRead) as BlockReference;
            if (sourceSample is null || targetSample is null)
            {
                _log.Write("Выбранный объект не является блоком.");
                return;
            }

            ObjectId sourceBtrId = sourceSample.BlockTableRecord;
            ObjectId targetBtrId = targetSample.BlockTableRecord;
            if (sourceBtrId == targetBtrId)
            {
                _log.Write("Исходный и целевой блок совпадают. Замена не требуется.");
                return;
            }

            var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            foreach (ObjectId btrId in blockTable)
            {
                var space = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                if (!space.IsLayout)
                {
                    continue;
                }

                var sourceRefs = new List<BlockReference>();
                foreach (ObjectId entId in space)
                {
                    var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                    if (br is not null && br.BlockTableRecord == sourceBtrId)
                    {
                        sourceRefs.Add(br);
                    }
                }

                if (sourceRefs.Count == 0)
                {
                    continue;
                }

                space.UpgradeOpen();
                foreach (BlockReference sourceRef in sourceRefs)
                {
                    var newRef = new BlockReference(sourceRef.Position, targetBtrId)
                    {
                        Rotation = sourceRef.Rotation,
                        ScaleFactors = sourceRef.ScaleFactors,
                        Layer = sourceRef.Layer
                    };

                    space.AppendEntity(newRef);
                    tr.AddNewlyCreatedDBObject(newRef, true);
                    sourceRef.UpgradeOpen();
                    sourceRef.Erase();
                    replaced++;
                }
            }
        });

        _log.Write($"Замена блоков завершена. Заменено: {replaced}.");
        return replaced;
        // END_BLOCK_EXECUTE_MAPPING
    }
}
