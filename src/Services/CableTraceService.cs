// FILE: src/Services/CableTraceService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Trace selected route and compute cable length including vertical delta and reserve.
//   SCOPE: Polyline lookup, height extraction, formula calculation, XData write.
//   DEPENDS: M-GEOM, M-ATTR, M-XDATA, M-LOG
//   LINKS: M-TRACE, M-GEOM, M-ATTR, M-XDATA, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   ExecuteTrace - Performs trace workflow and persists EOM_DATA payload.
// END_MODULE_MAP

using ElTools.Models;
using ElTools.Data;
using ElTools.Integrations;
using ElTools.Shared;

namespace ElTools.Services;

public class CableTraceService
{
    private readonly AutoCADAdapter _acad = new();
    private readonly GeometryService _geometry = new();
    private readonly AttributeService _attributes = new();
    private readonly XDataService _xdata = new();
    private readonly SettingsRepository _settings = new();
    private readonly IInstallTypeResolver _installTypeResolver = new InstallTypeResolver();
    private readonly LogService _log = new();
    private const string HeightTag = "ВЫСОТА_УСТАНОВКИ";

    public TraceResult? ExecuteTrace(TraceRequest request)
    {
        // START_BLOCK_EXECUTE_TRACE
        return ExecuteTraceFromBase(request.PolylineId, request.TargetBlockId);
        // END_BLOCK_EXECUTE_TRACE
    }

    public TraceResult? ExecuteTraceFromBase(ObjectId basePolylineId, ObjectId targetBlockId)
    {
        // START_BLOCK_EXECUTE_TRACE_FROM_BASE
        Document? doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (doc is null || basePolylineId.IsNull || targetBlockId.IsNull)
        {
            return null;
        }

        TraceResult? result = null;

        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            var basePolyline = tr.GetObject(basePolylineId, OpenMode.ForRead) as Polyline;
            var targetBlock = tr.GetObject(targetBlockId, OpenMode.ForRead) as BlockReference;
            if (basePolyline is null || targetBlock is null)
            {
                return null;
            }

            Point3d targetPoint = targetBlock.Position;
            Point3d closestOnBase = basePolyline.GetClosestPointTo(targetPoint, false);
            Point3d normalizedClosest = _geometry.NormalizeUcsToWcs([closestOnBase])[0];
            Point3d normalizedTarget = _geometry.NormalizeUcsToWcs([targetPoint])[0];

            var owner = tr.GetObject(basePolyline.OwnerId, OpenMode.ForWrite) as BlockTableRecord;
            if (owner is null)
            {
                return null;
            }

            Polyline branch = BuildOrthogonalBranchPolyline(basePolyline, normalizedClosest, normalizedTarget);
            owner.AppendEntity(branch);
            tr.AddNewlyCreatedDBObject(branch, true);

            double sourceHeight = basePolyline.StartPoint.Z;
            double targetHeight = ReadHeight(_attributes.ReadAttributes(targetBlockId));
            double total = _geometry.CalculateCableLength(branch, sourceHeight, targetHeight);

            result = new TraceResult(branch.Length, Math.Abs(sourceHeight - targetHeight), total, "N/A");
            tr.Commit();
        }

        _log.Write($"Трассировка завершена. Длина: {result?.TotalLength:0.###} м.");
        return result;
        // END_BLOCK_EXECUTE_TRACE_FROM_BASE
    }

    public IReadOnlyList<GroupTraceAggregate> RecalculateByGroups()
    {
        // START_BLOCK_RECALCULATE_BY_GROUPS
        Document? doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        if (doc is null)
        {
            return [];
        }

        _ = _settings.LoadSettings();
        InstallTypeRuleSet rules = _settings.LoadInstallTypeRules();
        var aggregates = new Dictionary<string, MutableAggregate>(StringComparer.OrdinalIgnoreCase);
        int defaultInstallTypeLines = 0;

        using (doc.LockDocument())
        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
            var blockTable = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in modelSpace)
            {
                if (tr.GetObject(id, OpenMode.ForRead) is not Entity entity)
                {
                    continue;
                }

                if (entity is BlockReference block)
                {
                    RegisterLoadBlock(aggregates, block.ObjectId);
                    continue;
                }

                if (entity is not (Line or Polyline))
                {
                    continue;
                }

                string? group = _xdata.GetLineGroup(id);
                if (string.IsNullOrWhiteSpace(group))
                {
                    continue;
                }

                string linetypeResolved = _acad.ResolveLinetype(entity);
                string installType = _installTypeResolver.Resolve(linetypeResolved, entity.Layer, rules);
                if (string.Equals(installType, rules.Default, StringComparison.OrdinalIgnoreCase))
                {
                    defaultInstallTypeLines++;
                }

                double length = entity switch
                {
                    Line line => line.Length,
                    Polyline polyline => polyline.Length,
                    _ => 0
                };

                string key = $"|{group.Trim()}";
                if (!aggregates.TryGetValue(key, out MutableAggregate? aggregate))
                {
                    aggregate = new MutableAggregate(string.Empty, group.Trim());
                    aggregates[key] = aggregate;
                }

                aggregate.TotalLength += Math.Round(length, 3);
                aggregate.LengthByInstallType.TryGetValue(installType, out double current);
                aggregate.LengthByInstallType[installType] = Math.Round(current + length, 3);
            }

            tr.Commit();
        }

        _log.Write($"[CableTraceService][RecalculateByGroups][RECALCULATE] Групп: {aggregates.Count}, линий Default: {defaultInstallTypeLines}.");
        return aggregates.Values
            .Select(a => new GroupTraceAggregate(
                a.Shield,
                a.Group,
                Math.Round(a.TotalPowerWatts, 3),
                Math.Round(a.TotalLength, 3),
                new Dictionary<string, double>(a.LengthByInstallType)))
            .OrderBy(a => a.Shield)
            .ThenBy(a => a.Group)
            .ToList();
        // END_BLOCK_RECALCULATE_BY_GROUPS
    }

    private static Polyline BuildOrthogonalBranchPolyline(Polyline basePolyline, Point3d splitPoint, Point3d targetPoint)
    {
        // START_BLOCK_BUILD_ORTHOGONAL_BRANCH_POLYLINE
        var branch = new Polyline();
        ApplyBasePolylineStyle(basePolyline, branch);

        double splitParam = basePolyline.GetParameterAtPoint(splitPoint);
        int splitWholeIndex = (int)Math.Floor(splitParam);
        bool splitOnVertex = Math.Abs(splitParam - Math.Round(splitParam)) < 1e-8;

        int vertexIndex = 0;
        for (int i = 0; i <= splitWholeIndex; i++)
        {
            Point3d v = basePolyline.GetPoint3dAt(i);
            branch.AddVertexAt(vertexIndex++, new Point2d(v.X, v.Y), 0, 0, 0);
        }

        if (!splitOnVertex)
        {
            branch.AddVertexAt(vertexIndex++, new Point2d(splitPoint.X, splitPoint.Y), 0, 0, 0);
        }

        Point2d orthoCorner = BuildOrthogonalCorner(splitPoint, targetPoint);
        if (!IsSamePoint2d(orthoCorner, new Point2d(splitPoint.X, splitPoint.Y)))
        {
            branch.AddVertexAt(vertexIndex++, orthoCorner, 0, 0, 0);
        }

        branch.AddVertexAt(vertexIndex, new Point2d(targetPoint.X, targetPoint.Y), 0, 0, 0);
        return branch;
        // END_BLOCK_BUILD_ORTHOGONAL_BRANCH_POLYLINE
    }

    private static void ApplyBasePolylineStyle(Polyline source, Polyline target)
    {
        // START_BLOCK_APPLY_BASE_POLYLINE_STYLE
        target.Layer = source.Layer;
        target.Linetype = source.Linetype;
        target.LinetypeScale = source.LinetypeScale;
        target.ConstantWidth = source.ConstantWidth;
        target.Color = source.Color;
        target.LineWeight = source.LineWeight;
        target.Transparency = source.Transparency;
        // END_BLOCK_APPLY_BASE_POLYLINE_STYLE
    }

    private static Point2d BuildOrthogonalCorner(Point3d splitPoint, Point3d targetPoint)
    {
        // START_BLOCK_BUILD_ORTHOGONAL_CORNER
        double dx = Math.Abs(targetPoint.X - splitPoint.X);
        double dy = Math.Abs(targetPoint.Y - splitPoint.Y);
        if (dx >= dy)
        {
            return new Point2d(targetPoint.X, splitPoint.Y);
        }

        return new Point2d(splitPoint.X, targetPoint.Y);
        // END_BLOCK_BUILD_ORTHOGONAL_CORNER
    }

    private static bool IsSamePoint2d(Point2d a, Point2d b)
    {
        // START_BLOCK_COMPARE_POINTS_2D
        const double eps = 1e-8;
        return Math.Abs(a.X - b.X) < eps && Math.Abs(a.Y - b.Y) < eps;
        // END_BLOCK_COMPARE_POINTS_2D
    }

    private static double ReadHeight(IReadOnlyDictionary<string, string> attributes)
    {
        // START_BLOCK_READ_HEIGHT
        if (!attributes.TryGetValue(HeightTag, out string? raw))
        {
            return 0;
        }

        return double.TryParse(raw, out double value) ? value : 0;
        // END_BLOCK_READ_HEIGHT
    }

    private void RegisterLoadBlock(IDictionary<string, MutableAggregate> aggregates, ObjectId blockId)
    {
        // START_BLOCK_REGISTER_LOAD_BLOCK
        IReadOnlyDictionary<string, string> attributes = _attributes.ReadAttributes(blockId);
        if (!attributes.TryGetValue(PluginConfig.AttributeTags.Group, out string? group) || string.IsNullOrWhiteSpace(group))
        {
            return;
        }

        if (!attributes.TryGetValue(PluginConfig.AttributeTags.Shield, out string? shield))
        {
            shield = string.Empty;
        }

        double power = 0;
        if (attributes.TryGetValue(PluginConfig.AttributeTags.Power, out string? powerRaw))
        {
            _ = double.TryParse(powerRaw?.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out power);
        }

        string key = $"{shield}|{group.Trim()}";
        if (!aggregates.TryGetValue(key, out MutableAggregate? aggregate))
        {
            aggregate = new MutableAggregate(shield?.Trim() ?? string.Empty, group.Trim());
            aggregates[key] = aggregate;
        }

        aggregate.TotalPowerWatts += Math.Max(power, 0);
        // END_BLOCK_REGISTER_LOAD_BLOCK
    }

    private sealed class MutableAggregate
    {
        public MutableAggregate(string shield, string group)
        {
            Shield = shield;
            Group = group;
        }

        public string Shield { get; }
        public string Group { get; }
        public double TotalPowerWatts { get; set; }
        public double TotalLength { get; set; }
        public Dictionary<string, double> LengthByInstallType { get; } = new(StringComparer.OrdinalIgnoreCase);
    }
}
