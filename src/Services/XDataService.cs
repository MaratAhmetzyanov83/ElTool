// FILE: src/Services/XDataService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Read and write structured XData with application key EOM_DATA.
//   SCOPE: Stores and retrieves plugin payload on AutoCAD entities.
//   DEPENDS: M-ACAD
//   LINKS: M-XDATA, M-ACAD
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   WriteEomData - Writes EOM_DATA payload to entity XData.
//   ReadEomData - Reads EOM_DATA payload from entity XData.
// END_MODULE_MAP

using ElTools.Integrations;
using ElTools.Models;

namespace ElTools.Services;

public class XDataService
{
    private const string AppName = "EOM_DATA";
    private readonly AutoCADAdapter _acad = new();

    public void WriteEomData(ObjectId entityId, EomDataRecord record)
    {
        // START_BLOCK_WRITE_EOM_DATA
        _acad.RunTransaction((tr, db) =>
        {
            var regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!regTable.Has(AppName))
            {
                regTable.UpgradeOpen();
                var reg = new RegAppTableRecord { Name = AppName };
                regTable.Add(reg);
                tr.AddNewlyCreatedDBObject(reg, true);
            }

            var entity = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;
            if (entity is null)
            {
                return;
            }

            entity.XData = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, AppName),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, record.CableMark),
                new TypedValue((int)DxfCode.ExtendedDataReal, record.TotalLength),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, record.Group),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, record.Type)
            );
        });
        // END_BLOCK_WRITE_EOM_DATA
    }

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
}
