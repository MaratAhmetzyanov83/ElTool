// FILE: src/Services/BlockMappingService.cs
// VERSION: 1.1.0
// START_MODULE_CONTRACT
//   PURPOSE: Replace source blocks with a selected target block in drawing.
//   SCOPE: Source/target template detection, in-place block replacement, replacement summary.
//   DEPENDS: M-CAD-CONTEXT, M-LOGGING
//   LINKS: M-BLOCK-REPLACE, M-CAD-CONTEXT, M-LOGGING
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

    // START_CONTRACT: ExecuteMapping
    //   PURPOSE: Execute mapping.
    //   INPUTS: { sourceSampleId: ObjectId - method parameter; targetSampleId: ObjectId - method parameter }
    //   OUTPUTS: { int - result of execute mapping }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-BLOCK-REPLACE
    // END_CONTRACT: ExecuteMapping

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
                _log.Write("Р вЂ™РЎвЂ№Р В±РЎР‚Р В°Р Р…Р Р…РЎвЂ№Р в„– Р С•Р В±РЎР‰Р ВµР С”РЎвЂљ Р Р…Р Вµ РЎРЏР Р†Р В»РЎРЏР ВµРЎвЂљРЎРѓРЎРЏ Р В±Р В»Р С•Р С”Р С•Р С.");
                return;
            }

            ObjectId sourceBtrId = GetEffectiveBlockDefinitionId(sourceSample);
            ObjectId targetBtrId = GetEffectiveBlockDefinitionId(targetSample);
            if (sourceBtrId == targetBtrId)
            {
                _log.Write("Р ВРЎРѓРЎвЂ¦Р С•Р Т‘Р Р…РЎвЂ№Р в„– Р С‘ РЎвЂ Р ВµР В»Р ВµР Р†Р С•Р в„– Р В±Р В»Р С•Р С” РЎРѓР С•Р Р†Р С—Р В°Р Т‘Р В°РЎР‹РЎвЂљ. Р вЂ”Р В°Р СР ВµР Р…Р В° Р Р…Р Вµ РЎвЂљРЎР‚Р ВµР В±РЎС“Р ВµРЎвЂљРЎРѓРЎРЏ.");
                return;
            }

            // Scale is calibrated by the selected sample pair so visual size is preserved.
            Scale3d sourceSampleScale = sourceSample.ScaleFactors;
            Scale3d targetSampleScale = targetSample.ScaleFactors;
            double mx = SafeScaleRatio(targetSampleScale.X, sourceSampleScale.X);
            double my = SafeScaleRatio(targetSampleScale.Y, sourceSampleScale.Y);
            double mz = SafeScaleRatio(targetSampleScale.Z, sourceSampleScale.Z);

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
                    if (br is not null && GetEffectiveBlockDefinitionId(br) == sourceBtrId)
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
                    Scale3d srcScale = sourceRef.ScaleFactors;
                    var newRef = new BlockReference(sourceRef.Position, targetBtrId)
                    {
                        Rotation = sourceRef.Rotation,
                        ScaleFactors = new Scale3d(srcScale.X * mx, srcScale.Y * my, srcScale.Z * mz),
                        Layer = sourceRef.Layer,
                        Normal = sourceRef.Normal
                    };

                    space.AppendEntity(newRef);
                    tr.AddNewlyCreatedDBObject(newRef, true);
                    sourceRef.UpgradeOpen();
                    sourceRef.Erase();
                    replaced++;
                }
            }
        });

        _log.Write($"Р вЂ”Р В°Р СР ВµР Р…Р В° Р В±Р В»Р С•Р С”Р С•Р Р† Р В·Р В°Р Р†Р ВµРЎР‚РЎв‚¬Р ВµР Р…Р В°. Р вЂ”Р В°Р СР ВµР Р…Р ВµР Р…Р С•: {replaced}.");
        return replaced;
        // END_BLOCK_EXECUTE_MAPPING
    }

    // START_CONTRACT: GetEffectiveBlockDefinitionId
    //   PURPOSE: Retrieve effective block definition id.
    //   INPUTS: { blockReference: BlockReference - method parameter }
    //   OUTPUTS: { ObjectId - result of retrieve effective block definition id }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-BLOCK-REPLACE
    // END_CONTRACT: GetEffectiveBlockDefinitionId

    private static ObjectId GetEffectiveBlockDefinitionId(BlockReference blockReference)
    {
        // START_BLOCK_GET_EFFECTIVE_BLOCK_DEFINITION_ID
        if (!blockReference.DynamicBlockTableRecord.IsNull)
        {
            return blockReference.DynamicBlockTableRecord;
        }

        return blockReference.BlockTableRecord;
        // END_BLOCK_GET_EFFECTIVE_BLOCK_DEFINITION_ID
    }

    // START_CONTRACT: SafeScaleRatio
    //   PURPOSE: Safe scale ratio.
    //   INPUTS: { target: double - method parameter; source: double - method parameter }
    //   OUTPUTS: { double - result of safe scale ratio }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-BLOCK-REPLACE
    // END_CONTRACT: SafeScaleRatio

    private static double SafeScaleRatio(double target, double source)
    {
        // START_BLOCK_SAFE_SCALE_RATIO
        if (Math.Abs(source) < 1e-9)
        {
            return 1.0;
        }

        return target / source;
        // END_BLOCK_SAFE_SCALE_RATIO
    }
}