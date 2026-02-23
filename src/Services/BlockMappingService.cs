// FILE: src/Services/BlockMappingService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Replace designer blocks by organization-standard blocks using mapping rules from settings.
//   SCOPE: Block name matching, attribute transfer, replacement summary.
//   DEPENDS: M-SETTINGS, M-ATTR, M-ACAD, M-LOG
//   LINKS: M-MAP, M-SETTINGS, M-ATTR, M-ACAD, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   ExecuteMapping - Applies mapping rules for selected block references.
// END_MODULE_MAP

using ElTools.Data;
using ElTools.Integrations;
using ElTools.Models;

namespace ElTools.Services;

public class BlockMappingService
{
    private readonly SettingsRepository _settings = new();
    private readonly AttributeService _attributes = new();
    private readonly AutoCADAdapter _acad = new();
    private readonly LogService _log = new();

    public int ExecuteMapping(IEnumerable<ObjectId> blockIds)
    {
        // START_BLOCK_EXECUTE_MAPPING
        SettingsModel settings = _settings.LoadSettings();
        MappingProfile? profile = settings.MappingProfiles.FirstOrDefault(x => x.Name == settings.ActiveProfile)
            ?? settings.MappingProfiles.FirstOrDefault();

        if (profile is null)
        {
            _log.Write("Профиль маппинга отсутствует.");
            return 0;
        }

        int replaced = 0;
        foreach (ObjectId blockId in blockIds)
        {
            _acad.RunTransaction((tr, db) =>
            {
                var sourceRef = tr.GetObject(blockId, OpenMode.ForRead) as BlockReference;
                if (sourceRef is null)
                {
                    return;
                }

                string sourceName = string.Empty;
                if (tr.GetObject(sourceRef.BlockTableRecord, OpenMode.ForRead) is BlockTableRecord sourceBtr)
                {
                    sourceName = sourceBtr.Name;
                }

                MappingRule? rule = profile.Rules.FirstOrDefault(r => string.Equals(r.SourceBlockName, sourceName, StringComparison.OrdinalIgnoreCase));
                if (rule is null)
                {
                    return;
                }

                var blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!blockTable.Has(rule.TargetBlockName))
                {
                    return;
                }

                var owner = (BlockTableRecord)tr.GetObject(sourceRef.OwnerId, OpenMode.ForWrite);
                var targetBtrId = blockTable[rule.TargetBlockName];

                var newRef = new BlockReference(sourceRef.Position, targetBtrId)
                {
                    Rotation = sourceRef.Rotation,
                    ScaleFactors = sourceRef.ScaleFactors
                };

                owner.AppendEntity(newRef);
                tr.AddNewlyCreatedDBObject(newRef, true);

                Dictionary<string, string> srcAttributes = _attributes.ReadAttributes(blockId);
                _attributes.WriteAttributes(newRef.ObjectId, srcAttributes);

                sourceRef.UpgradeOpen();
                sourceRef.Erase();
                replaced++;
            });
        }

        _log.Write($"Замена блоков завершена. Заменено: {replaced}.");
        return replaced;
        // END_BLOCK_EXECUTE_MAPPING
    }
}
