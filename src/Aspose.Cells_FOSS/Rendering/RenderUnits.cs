using System;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Central place for the unit conversions used when leaving Excel's native units. All layout
    /// coordinates downstream of this type are expressed in PDF points (1pt = 1/72 inch).
    /// </summary>
    internal static class RenderUnits
    {
        /// <summary>Points per inch (PDF user-space unit is the point).</summary>
        public const double PointsPerInch = 72d;

        /// <summary>Excel measures pixels at 96 DPI; a pixel is therefore 0.75 points.</summary>
        public const double PixelsToPoints = 72d / 96d;

        /// <summary>Inverse of <see cref="PixelsToPoints"/>: points expressed in 96-DPI pixels.</summary>
        public const double PointsToPixels = 96d / 72d;

        public static double InchesToPoints(double inches)
        {
            return inches * PointsPerInch;
        }

        public static double PixelsToPointsValue(double pixels)
        {
            return pixels * PixelsToPoints;
        }

        /// <summary>
        /// Converts an Excel column width (measured in characters of the digit '0') to pixels using
        /// the workbook's Maximum Digit Width, per ECMA-376.
        /// </summary>
        public static double ColumnWidthCharsToPixels(double widthChars, double maxDigitWidth)
        {
            if (maxDigitWidth <= 0d)
            {
                maxDigitWidth = 7.5d;
            }

            // The ECMA formula truncates to whole device pixels for screen rendering; for a vector
            // PDF we keep the sub-pixel value so column widths match Excel's export precisely. The
            // inner padding term (128/MDW) is retained per the spec.
            return (256d * widthChars + Math.Truncate(128d / maxDigitWidth)) / 256d * maxDigitWidth;
        }

        public static double ColumnWidthCharsToPoints(double widthChars, double maxDigitWidth)
        {
            return PixelsToPointsValue(ColumnWidthCharsToPixels(widthChars, maxDigitWidth));
        }

        /// <summary>
        /// Converts Excel's default column width, which is expressed as a count of display characters
        /// (8.43 by default) rather than a stored/padded width, to points: pixels = round(chars * MDW) +
        /// cell padding. This is the display-side inverse of <see cref="ColumnWidthCharsToPixels"/>,
        /// which expects an already padded stored width (used for explicit <c>&lt;col&gt;</c> widths).
        /// The 4px padding (2px per side) matches Excel's measured default-column export width.
        /// </summary>
        public static double DefaultColumnWidthCharsToPoints(double displayChars, double maxDigitWidth)
        {
            if (maxDigitWidth <= 0d)
            {
                maxDigitWidth = 7.5d;
            }

            var pixels = Math.Round(displayChars * maxDigitWidth, MidpointRounding.AwayFromZero) + 4d;
            return PixelsToPointsValue(pixels);
        }
    }

    /// <summary>
    /// Maps an OOXML paper-size code to physical dimensions in points (portrait orientation).
    /// </summary>
    internal static class PaperSizes
    {
        public static void GetDimensionsPoints(int paperSizeCode, out double widthPt, out double heightPt)
        {
            // Dimensions are in inches, converted to points. When the workbook does not specify a
            // paper size (code 0), Excel falls back to the machine/regional default; we default to
            // A4 to match the common non-US default and the reference exports we validate against.
            double widthIn, heightIn;
            switch (paperSizeCode)
            {
                case 1:  widthIn = 8.5;  heightIn = 11.0;  break; // Letter
                case 3:  widthIn = 11.0; heightIn = 17.0;  break; // Tabloid
                case 5:  widthIn = 8.5;  heightIn = 14.0;  break; // Legal
                case 8:  widthIn = 11.69; heightIn = 16.54; break; // A3
                case 9:  widthIn = 8.27; heightIn = 11.69; break; // A4
                case 11: widthIn = 5.83; heightIn = 8.27;  break; // A5
                case 13: widthIn = 7.17; heightIn = 10.12; break; // B5 (JIS)
                default: widthIn = 8.27; heightIn = 11.69; break; // A4 fallback (unspecified)
            }

            widthPt = RenderUnits.InchesToPoints(widthIn);
            heightPt = RenderUnits.InchesToPoints(heightIn);
        }
    }
}
