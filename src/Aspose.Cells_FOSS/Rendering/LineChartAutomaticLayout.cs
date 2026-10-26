using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class LineChartAutomaticLayout
    {
        public static bool Applies(ParsedChart chart)
        {
            return IsAutomaticLine(chart) && chart.HasDateCategoryAxis;
        }

        public static bool AppliesToCategoryAxis(ParsedChart chart)
        {
            return IsAutomaticLine(chart) && !chart.HasDateCategoryAxis;
        }

        private static bool IsAutomaticLine(ParsedChart chart)
        {
            return chart != null && chart.Kind == ChartKind.Line && !chart.IsLine3D
                && chart.CategoryAxisAutomaticTextRotation;
        }

        public static SKRect Plot(SKRect area, ParsedChart chart, float labelWidth,
            float rotation, SKRect fallback)
        {
            var horizontalCategories = AppliesToCategoryAxis(chart) && Math.Abs(rotation) < 0.1f;
            if ((!Applies(chart) && !horizontalCategories) || chart.HasManualPlot
                || (!horizontalCategories && Math.Abs(rotation) < 0.1f)
                || chart.Title == null || !string.Equals(chart.LegendPosition, "b", StringComparison.Ordinal))
            {
                return fallback;
            }
            // Text bands stop growing once the fixed-size title, labels and legend fit.
            // Taller charts should give their extra height to the inner plot.
            var topReserve = Math.Min(area.Height * 0.185f, 36f);
            var bottomReserve = horizontalCategories
                ? Math.Min(area.Height * 0.239f, 46.5f)
                : Math.Min(area.Height * 0.318f, 62f);
            return new SKRect(area.Left + Math.Max(area.Width * 0.06f, labelWidth + 14.5f),
                area.Top + topReserve, area.Right - area.Width * 0.031f,
                area.Bottom - bottomReserve);
        }
    }
}
