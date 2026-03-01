// FILE: src/Services/GeometryService.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Compute lengths and geometry metrics with mandatory UCS to WCS normalization.
//   SCOPE: UCS->WCS route conversion and cable length formula.
//   DEPENDS: M-CAD-CONTEXT
//   LINKS: M-COORDINATES, M-CAD-CONTEXT
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   NormalizeUcsToWcs - Converts polyline points from UCS to WCS.
//   CalculateCableLength - Calculates L_total = L_2d + ABS(H1-H2) + 5%.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using ElTools.Integrations;

namespace ElTools.Services;

public class GeometryService
{
    private readonly AutoCADAdapter _acad = new();

    // START_CONTRACT: NormalizeUcsToWcs
    //   PURPOSE: Normalize ucs to wcs.
    //   INPUTS: { points: IEnumerable<Point3d> - method parameter }
    //   OUTPUTS: { IReadOnlyList<Point3d> - result of normalize ucs to wcs }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-COORDINATES
    // END_CONTRACT: NormalizeUcsToWcs

    public IReadOnlyList<Point3d> NormalizeUcsToWcs(IEnumerable<Point3d> points)
    {
        // START_BLOCK_NORMALIZE_UCS_TO_WCS
        return points.Select(_acad.UcsToWcs).ToList();
        // END_BLOCK_NORMALIZE_UCS_TO_WCS
    }

    // START_CONTRACT: CalculateCableLength
    //   PURPOSE: Calculate cable length.
    //   INPUTS: { polyline: Polyline - method parameter; sourceHeight: double - method parameter; targetHeight: double - method parameter }
    //   OUTPUTS: { double - result of calculate cable length }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-COORDINATES
    // END_CONTRACT: CalculateCableLength

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