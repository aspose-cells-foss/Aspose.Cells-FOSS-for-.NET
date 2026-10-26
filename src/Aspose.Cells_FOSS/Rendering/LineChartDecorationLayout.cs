using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class LineChartDecorationLayout
    {
        private static bool UsesManualLinePlot(ParsedChart chart)
        {
            return chart != null && chart.Kind == ChartKind.Line
                && !chart.IsLine3D && chart.HasManualPlot;
        }

        public static float TitleCenterY(SKRect area, ParsedChart chart, float fallback)
        {
            if (chart != null && chart.IsLine3D && !chart.HasDateCategoryAxis && !chart.HasManualTitle)
            {
                return area.Top + (chart.TitleFont != null ? (float)chart.TitleFont.Size : 14f) * 1.05f;
            }
            if ((!UsesManualLinePlot(chart) && !LineChartAutomaticLayout.Applies(chart)
                && !LineChartAutomaticLayout.AppliesToCategoryAxis(chart)) || chart.HasManualTitle)
            {
                return fallback;
            }
            // Keep the automatic title in its chart band, independent of the inner plot.
            var size = chart.TitleFont != null ? (float)chart.TitleFont.Size : 14f;
            return area.Top + size * 1.05f;
        }

        public static float LegendCenterY(SKRect area, ParsedChart chart, float fallback)
        {
            if (chart != null && chart.IsLine3D && !chart.HasDateCategoryAxis && !chart.HasManualLegend
                && string.Equals(chart.LegendPosition, "b", StringComparison.Ordinal))
            {
                return area.Bottom - 15.7f;
            }
            if (!UsesAutomaticBottomLegend(chart))
            {
                return fallback;
            }
            // Anchor the automatic legend to the chart, not the remaining plot gap.
            return area.Bottom - (chart.HasManualPlot ? 15.2f : 14.8f);
        }

        public static bool UsesAutomaticBottomLegend(ParsedChart chart)
        {
            return (UsesManualLinePlot(chart) || LineChartAutomaticLayout.Applies(chart)
                || LineChartAutomaticLayout.AppliesToCategoryAxis(chart)) && !chart.HasManualLegend
                && string.Equals(chart.LegendPosition, "b", StringComparison.Ordinal);
        }

        public static float SwatchWidth(ParsedChart chart, float fallback)
        {
            return UsesAutomaticBottomLegend(chart) ? 19.2f : fallback;
        }

        public static float TextGap(ParsedChart chart, float fallback)
        {
            return UsesAutomaticBottomLegend(chart) ? 2f : fallback;
        }

        public static float ItemGap(ParsedChart chart, float fallback)
        {
            if (UsesAutomaticBottomLegend(chart) && LineChartAutomaticLayout.AppliesToCategoryAxis(chart))
            {
                return 8.5f;
            }
            return UsesAutomaticBottomLegend(chart) ? 9.5f : fallback;
        }

        public static float HorizontalOffset(ParsedChart chart, float fallback)
        {
            // Match the optical centering of the line-key and text group.
            if (UsesAutomaticBottomLegend(chart) && LineChartAutomaticLayout.AppliesToCategoryAxis(chart))
            {
                return 2.25f;
            }
            return UsesAutomaticBottomLegend(chart) && !chart.HasManualPlot
                && LineChartAutomaticLayout.Applies(chart)
                ? 2.1f : fallback;
        }
    }
}
