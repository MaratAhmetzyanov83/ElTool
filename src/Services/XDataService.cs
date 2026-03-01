// FILE: src/Services/XDataService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Read and write structured XData with application key EOM_DATA.
//   SCOPE: Stores and retrieves plugin payload on AutoCAD entities.
//   DEPENDS: M-CAD-CONTEXT, M-CONFIG, M-MODELS
//   LINKS: M-XDATA, M-CAD-CONTEXT, M-CONFIG, M-MODELS
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   WriteEomData - Writes EOM_DATA payload to entity XData.
//   ReadEomData - Reads EOM_DATA payload from entity XData.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using ElTools.Integrations;
using ElTools.Models;
using ElTools.Shared;

namespace ElTools.Services;

public class XDataService
{
    private static readonly Dictionary<string, string> ActiveGroupsByDocument = new(StringComparer.OrdinalIgnoreCase);
    private const string LineMetadataSeparator = "|";
    private readonly AutoCADAdapter _acad = new();

    // START_CONTRACT: WriteEomData
    //   PURPOSE: Write eom data.
    //   INPUTS: { entityId: ObjectId - method parameter; record: EomDataRecord - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: WriteEomData

    public void WriteEomData(ObjectId entityId, EomDataRecord record)
    {
        // START_BLOCK_WRITE_EOM_DATA
        _acad.RunTransaction((tr, db) =>
        {
            EnsureRegApp(tr, db, PluginConfig.Metadata.EomDataAppName);

            var entity = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;
            if (entity is null)
            {
                return;
            }

            UpsertAppXData(entity, PluginConfig.Metadata.EomDataAppName, new[]
            {
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, record.CableMark),
                new TypedValue((int)DxfCode.ExtendedDataReal, record.TotalLength),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, record.Group),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, record.Type)
            });
        });
        // END_BLOCK_WRITE_EOM_DATA
    }

    // START_CONTRACT: ReadEomData
    //   PURPOSE: Read eom data.
    //   INPUTS: { entityId: ObjectId - method parameter }
    //   OUTPUTS: { EomDataRecord? - result of read eom data }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: ReadEomData

    public EomDataRecord? ReadEomData(ObjectId entityId)
    {
        // START_BLOCK_READ_EOM_DATA
        EomDataRecord? result = null;

        _acad.RunTransaction((tr, _) =>
        {
            var entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
            if (entity?.XData is null)
            {
                return;
            }

            TypedValue[] values = entity.XData.AsArray();
            if (values.Length < 5)
            {
                return;
            }

            result = new EomDataRecord(
                values[1].Value?.ToString() ?? string.Empty,
                values[2].Value is double len ? len : 0,
                values[3].Value?.ToString() ?? string.Empty,
                values[4].Value?.ToString() ?? string.Empty
            );
        });

        return result;
        // END_BLOCK_READ_EOM_DATA
    }

    // START_CONTRACT: SetActiveGroup
    //   PURPOSE: Set active group.
    //   INPUTS: { groupCode: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: SetActiveGroup

    public void SetActiveGroup(string groupCode)
    {
        // START_BLOCK_SET_ACTIVE_GROUP
        Document? doc = _acad.GetActiveDocument();
        if (doc is null || string.IsNullOrWhiteSpace(groupCode))
        {
            return;
        }

        ActiveGroupsByDocument[GetDocumentKey(doc)] = groupCode.Trim();
        // END_BLOCK_SET_ACTIVE_GROUP
    }

    // START_CONTRACT: GetActiveGroup
    //   PURPOSE: Retrieve active group.
    //   INPUTS: none
    //   OUTPUTS: { string? - result of retrieve active group }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: GetActiveGroup

    public string? GetActiveGroup()
    {
        // START_BLOCK_GET_ACTIVE_GROUP
        Document? doc = _acad.GetActiveDocument();
        if (doc is null)
        {
            return null;
        }

        return ActiveGroupsByDocument.TryGetValue(GetDocumentKey(doc), out string? group) ? group : null;
        // END_BLOCK_GET_ACTIVE_GROUP
    }

    // START_CONTRACT: ApplyActiveGroupToEntity
    //   PURPOSE: Apply active group to entity.
    //   INPUTS: { entityId: ObjectId - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: ApplyActiveGroupToEntity

    public void ApplyActiveGroupToEntity(ObjectId entityId)
    {
        // START_BLOCK_APPLY_ACTIVE_GROUP_TO_ENTITY
        string? group = GetActiveGroup();
        if (string.IsNullOrWhiteSpace(group))
        {
            return;
        }

        SetLineGroup(entityId, group);
        // END_BLOCK_APPLY_ACTIVE_GROUP_TO_ENTITY
    }

    // START_CONTRACT: AssignGroupToSelection
    //   PURPOSE: Assign group to selection.
    //   INPUTS: { entityIds: IEnumerable<ObjectId> - method parameter; groupCode: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: AssignGroupToSelection

    public void AssignGroupToSelection(IEnumerable<ObjectId> entityIds, string groupCode)
    {
        // START_BLOCK_ASSIGN_GROUP_TO_SELECTION
        if (string.IsNullOrWhiteSpace(groupCode))
        {
            return;
        }

        foreach (ObjectId entityId in entityIds)
        {
            SetLineGroup(entityId, groupCode.Trim());
        }
        // END_BLOCK_ASSIGN_GROUP_TO_SELECTION
    }

    // START_CONTRACT: SetLineGroup
    //   PURPOSE: Set line group.
    //   INPUTS: { entityId: ObjectId - method parameter; groupCode: string - method parameter; installType: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: SetLineGroup

    public void SetLineGroup(ObjectId entityId, string groupCode, string installType = "")
    {
        // START_BLOCK_SET_LINE_GROUP
        _acad.RunTransaction((tr, db) =>
        {
            EnsureRegApp(tr, db, PluginConfig.Metadata.ElToolAppName);
            var entity = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;
            if (entity is null || !IsCableLine(entity))
            {
                return;
            }

            string payload = $"{groupCode.Trim()}{LineMetadataSeparator}{installType.Trim()}";
            UpsertAppXData(entity, PluginConfig.Metadata.ElToolAppName, [
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, payload)
            ]);
        });
        // END_BLOCK_SET_LINE_GROUP
    }

    // START_CONTRACT: GetLineGroup
    //   PURPOSE: Retrieve line group.
    //   INPUTS: { entityId: ObjectId - method parameter }
    //   OUTPUTS: { string? - result of retrieve line group }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: GetLineGroup

    public string? GetLineGroup(ObjectId entityId)
    {
        // START_BLOCK_GET_LINE_GROUP
        string? result = null;
        _acad.RunTransaction((tr, _) =>
        {
            var entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
            if (entity?.XData is null)
            {
                return;
            }

            TypedValue[] values = entity.XData.AsArray();
            if (values.Length < 2)
            {
                return;
            }

            if (!string.Equals(values[0].Value?.ToString(), PluginConfig.Metadata.ElToolAppName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string payload = values[1].Value?.ToString() ?? string.Empty;
            string[] parts = payload.Split(LineMetadataSeparator);
            result = parts.Length > 0 ? parts[0] : null;
        });

        return string.IsNullOrWhiteSpace(result) ? null : result;
        // END_BLOCK_GET_LINE_GROUP
    }

    // START_CONTRACT: IsCableLine
    //   PURPOSE: Check whether cable line.
    //   INPUTS: { entity: Entity - method parameter }
    //   OUTPUTS: { bool - true when method can check whether cable line }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: IsCableLine

    private static bool IsCableLine(Entity entity)
    {
        // START_BLOCK_IS_CABLE_LINE
        return entity is Line or Polyline;
        // END_BLOCK_IS_CABLE_LINE
    }

    // START_CONTRACT: GetDocumentKey
    //   PURPOSE: Retrieve document key.
    //   INPUTS: { doc: Document - method parameter }
    //   OUTPUTS: { string - textual result for retrieve document key }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: GetDocumentKey

    private static string GetDocumentKey(Document doc)
    {
        // START_BLOCK_GET_DOCUMENT_KEY
        return string.IsNullOrWhiteSpace(doc.Name) ? doc.Database.FingerprintGuid.ToString() : doc.Name;
        // END_BLOCK_GET_DOCUMENT_KEY
    }

    // START_CONTRACT: EnsureRegApp
    //   PURPOSE: Ensure reg app.
    //   INPUTS: { tr: Transaction - method parameter; db: Database - method parameter; appName: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: EnsureRegApp

    private static void EnsureRegApp(Transaction tr, Database db, string appName)
    {
        // START_BLOCK_ENSURE_REGAPP
        var regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
        if (regTable.Has(appName))
        {
            return;
        }

        regTable.UpgradeOpen();
        var reg = new RegAppTableRecord { Name = appName };
        regTable.Add(reg);
        tr.AddNewlyCreatedDBObject(reg, true);
        // END_BLOCK_ENSURE_REGAPP
    }

    // START_CONTRACT: UpsertAppXData
    //   PURPOSE: Upsert app X data.
    //   INPUTS: { entity: Entity - method parameter; appName: string - method parameter; appPayload: IReadOnlyList<TypedValue> - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-XDATA
    // END_CONTRACT: UpsertAppXData

    private static void UpsertAppXData(Entity entity, string appName, IReadOnlyList<TypedValue> appPayload)
    {
        // START_BLOCK_UPSERT_APP_XDATA
        var allValues = entity.XData?.AsArray().ToList() ?? [];
        var result = new List<TypedValue>();
        bool skip = false;

        foreach (TypedValue value in allValues)
        {
            if (value.TypeCode == (int)DxfCode.ExtendedDataRegAppName)
            {
                string currentApp = value.Value?.ToString() ?? string.Empty;
                skip = string.Equals(currentApp, appName, StringComparison.OrdinalIgnoreCase);
                if (!skip)
                {
                    result.Add(value);
                }

                continue;
            }

            if (!skip)
            {
                result.Add(value);
            }
        }

        result.Add(new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName));
        result.AddRange(appPayload);
        entity.XData = new ResultBuffer(result.ToArray());
        // END_BLOCK_UPSERT_APP_XDATA
    }
}