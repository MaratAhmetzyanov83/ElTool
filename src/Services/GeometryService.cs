// FILE: src/Services/GeometryService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Compute lengths and geometry metrics with mandatory UCS to WCS normalization.
//   SCOPE: UCS->WCS route conversion and cable length formula.
//   DEPENDS: M-ACAD
//   LINKS: M-GEOM, M-ACAD
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   NormalizeUcsToWcs - Converts polyline points from UCS to WCS.
//   CalculateCableLength - Calculates L_total = L_2d + ABS(H1-H2) + 5%.
// END_MODULE_MAP

using ElTools.Integrations;

namespace ElTools.Services;

public class GeometryService
{
    private readonly AutoCADAdapter _acad = new();

    public IReadOnlyList<Point3d> NormalizeUcsToWcs(IEnumerable<Point3d> points)
    {
        // START_BLOCK_NORMALIZE_UCS_TO_WCS
        return points.Select(_acad.UcsToWcs).ToList();
        // END_BLOCK_NORMALIZE_UCS_TO_WCS
    }

    public double CalculateCableLength(Polyline polyline, double sourceHeight, double targetHeight)
    {
        // START_BLOCK_CALCULATE_CABLE_LENGTH
        double length2D = polyline.Length;
        double heightDelta = Math.Abs(sourceHeight - targetHeight);
        double total = (length2D + heightDelta) * 1.05;
        return Math.Round(total, 3);
        // END_BLOCK_CALCULATE_CABLE_LENGTH
    }
}
