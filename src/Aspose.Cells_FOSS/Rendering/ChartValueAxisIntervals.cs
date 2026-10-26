using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartValueAxisIntervals
    {
        public static double ResolveMajor(ParsedChart chart, SKRect plot, double major)
        {
            if (chart != null && chart.Kind == ChartKind.Scatter && !chart.MajorUnit.HasValue
                && plot.Height < 105f)
            {
                return major * 2d;
            }

            if (chart == null || !chart.IsPercentStacked || chart.MajorUnit.HasValue || major < 0.1d)
            {
                return major;
            }

            if (chart.HasDataTable)
            {
                return 0.5d;
            }

            // Ten intervals need at least 12 pt each for the 9 pt axis labels.
            // Compact 2D line plots use five intervals; taller plots retain ten.
            if (chart.Kind == ChartKind.Line && !chart.IsLine3D && plot.Height < 120f)
            {
                return 0.2d;
            }

            return 0.1d;
        }
    }
}
