// FILE: src/Services/SpecificationService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Scan drawing EOM_DATA and aggregate cable lengths by type and group.
//   SCOPE: Reads XData records and groups by type/group.
//   DEPENDS: M-XDATA, M-SETTINGS, M-LOG
//   LINKS: M-SPEC, M-XDATA, M-SETTINGS, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   BuildSpecification - Produces grouped specification rows.
// END_MODULE_MAP

using ElTools.Data;
using ElTools.Models;

namespace ElTools.Services;

public class SpecificationService
{
    private readonly XDataService _xdata = new();
    private readonly SettingsRepository _settings = new();
    private readonly LogService _log = new();

    public IReadOnlyList<SpecificationRow> BuildSpecification(IEnumerable<ObjectId> entityIds)
    {
        // START_BLOCK_BUILD_SPECIFICATION
        _ = _settings.LoadSettings();

        var records = new List<EomDataRecord>();
        foreach (ObjectId id in entityIds)
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

        _log.Write($"Спецификация собрана. Строк: {grouped.Count}.");
        return grouped;
        // END_BLOCK_BUILD_SPECIFICATION
    }
}
