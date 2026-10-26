using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Excel-compatible print placement for worksheets containing ChartEx region maps.</summary>
    internal static class ChartExRegionMapPrintLayout
    {
        public const double Scale = 0.95d;
        public const double LeftInsetPt = 2.5d;
        public const double HorizontalBreakAllowancePt = 10d;

        public static bool Applies(WorksheetModel sheet)
        {
            if (sheet == null)
            {
                return false;
            }

            for (var i = 0; i < sheet.Charts.Count; i++)
            {
                var chart = sheet.Charts[i];
                if (chart != null && chart.IsChartEx && chart.RawChartXml != null
                    && chart.RawChartXml.IndexOf("regionMap", StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static double TopOrigin(double pageHeightPt, double topMarginPt)
        {
            // The Excel map-chart print reference places the sheet about 100pt from the top of A4.
            return Math.Max(topMarginPt, pageHeightPt * (100d / 842d));
        }
    }
}
