using Autodesk.Revit.DB;

namespace BOQExporter.Revit.Extraction
{
    internal static class QuantityUtils
    {
        /// <summary>
        /// Converts Revit internal area units (ft²) to square meters.
        /// </summary>
        public static double ToSquareMeters(double internalArea)
        {
            return UnitUtils.ConvertFromInternalUnits(internalArea, UnitTypeId.SquareMeters);
        }

        /// <summary>
        /// Converts Revit internal volume units (ft³) to cubic meters.
        /// </summary>
        public static double ToCubicMeters(double internalVolume)
        {
            return UnitUtils.ConvertFromInternalUnits(internalVolume, UnitTypeId.CubicMeters);
        }

        /// <summary>
        /// Converts Revit internal length units (ft) to meters.
        /// </summary>
        public static double ToMeters(double internalLength)
        {
            return UnitUtils.ConvertFromInternalUnits(internalLength, UnitTypeId.Meters);
        }
    }
}
