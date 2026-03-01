// FILE: src/Services/CableTraceService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Trace selected route and compute cable length including vertical delta and reserve.
//   SCOPE: Polyline lookup, height extraction, formula calculation, XData write.
//   DEPENDS: M-COORDINATES, M-ATTRIBUTES, M-XDATA, M-LOGGING, M-SETTINGS, M-INSTALL-TYPE-RESOLVER
//   LINKS: M-CABLE-CALC, M-COORDINATES, M-ATTRIBUTES, M-XDATA, M-LOGGING, M-SETTINGS, M-INSTALL-TYPE-RESOLVER
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   ExecuteTrace - Performs trace workflow and persists EOM_DATA payload.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


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
    private const string HeightTag = "Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В«Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ў_Р В Р’В Р В РІвЂљВ¬Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р РЋРІвЂћСћР В Р’В Р вЂ™Р’В";

    // START_CONTRACT: ExecuteTrace
    //   PURPOSE: Execute trace.
    //   INPUTS: { request: TraceRequest - method parameter }
    //   OUTPUTS: { TraceResult? - result of execute trace }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: ExecuteTrace

    public TraceResult? ExecuteTrace(TraceRequest request)
    {
        // START_BLOCK_EXECUTE_TRACE
        return ExecuteTraceFromBase(request.PolylineId, request.TargetBlockId);
        // END_BLOCK_EXECUTE_TRACE
    }

    // START_CONTRACT: ExecuteTraceFromBase
    //   PURPOSE: Execute trace from base.
    //   INPUTS: { basePolylineId: ObjectId - method parameter; targetBlockId: ObjectId - method parameter }
    //   OUTPUTS: { TraceResult? - result of execute trace from base }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: ExecuteTraceFromBase

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

        _log.Write($"Р В Р’В Р РЋРЎвЂєР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р Р‹Р В РЎвЂњР В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚ВР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚СњР В Р’В Р вЂ™Р’В° Р В Р’В Р вЂ™Р’В·Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РІР‚С™Р В Р Р‹Р Р†РІР‚С™Р’В¬Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’В°. Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’В°: {result?.TotalLength:0.###} Р В Р’В Р РЋР’В.");
        return result;
        // END_BLOCK_EXECUTE_TRACE_FROM_BASE
    }

    // START_CONTRACT: RecalculateByGroups
    //   PURPOSE: Recalculate by groups.
    //   INPUTS: none
    //   OUTPUTS: { IReadOnlyList<GroupTraceAggregate> - result of recalculate by groups }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: RecalculateByGroups

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

        _log.Write($"[CableTraceService][RecalculateByGroups][RECALCULATE] Р В Р’В Р Р†Р вЂљРЎС™Р В Р Р‹Р В РІР‚С™Р В Р Р‹Р РЋРІР‚СљР В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚вЂќ: {aggregates.Count}, Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р’В Р Р†РІР‚С›РІР‚вЂњ Default: {defaultInstallTypeLines}.");
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

    // START_CONTRACT: BuildOrthogonalBranchPolyline
    //   PURPOSE: Build orthogonal branch polyline.
    //   INPUTS: { basePolyline: Polyline - method parameter; splitPoint: Point3d - method parameter; targetPoint: Point3d - method parameter }
    //   OUTPUTS: { Polyline - result of build orthogonal branch polyline }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: BuildOrthogonalBranchPolyline

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

    // START_CONTRACT: ApplyBasePolylineStyle
    //   PURPOSE: Apply base polyline style.
    //   INPUTS: { source: Polyline - method parameter; target: Polyline - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: ApplyBasePolylineStyle

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

    // START_CONTRACT: BuildOrthogonalCorner
    //   PURPOSE: Build orthogonal corner.
    //   INPUTS: { splitPoint: Point3d - method parameter; targetPoint: Point3d - method parameter }
    //   OUTPUTS: { Point2d - result of build orthogonal corner }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: BuildOrthogonalCorner

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

    // START_CONTRACT: IsSamePoint2d
    //   PURPOSE: Check whether same point2d.
    //   INPUTS: { a: Point2d - method parameter; b: Point2d - method parameter }
    //   OUTPUTS: { bool - true when method can check whether same point2d }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: IsSamePoint2d

    private static bool IsSamePoint2d(Point2d a, Point2d b)
    {
        // START_BLOCK_COMPARE_POINTS_2D
        const double eps = 1e-8;
        return Math.Abs(a.X - b.X) < eps && Math.Abs(a.Y - b.Y) < eps;
        // END_BLOCK_COMPARE_POINTS_2D
    }

    // START_CONTRACT: ReadHeight
    //   PURPOSE: Read height.
    //   INPUTS: { attributes: IReadOnlyDictionary<string, string> - method parameter }
    //   OUTPUTS: { double - result of read height }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: ReadHeight

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

    // START_CONTRACT: RegisterLoadBlock
    //   PURPOSE: Register load block.
    //   INPUTS: { aggregates: IDictionary<string, MutableAggregate> - method parameter; blockId: ObjectId - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-CABLE-CALC
    // END_CONTRACT: RegisterLoadBlock

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