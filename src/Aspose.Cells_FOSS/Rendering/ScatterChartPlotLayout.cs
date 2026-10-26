using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ScatterChartPlotLayout
    {
        public static SKRect Resolve(SKRect area, ParsedChart chart)
        {
            var dateAxis = ChartXmlParser.IsDateFormat(chart.ScatterXFormatCode);
            var rightReserve = chart.LegendPosition == null ? 0.05f
                : dateAxis ? 0.06f : 0.078f;
            var leftReserve = chart.LegendPosition == null ? 0.10f
                : string.IsNullOrEmpty(chart.ValueDisplayUnitLabel) ? 0.132f : 0.162f;
            var bottomReserve = chart.LegendPosition == null ? 0.11f : 0.235f;
            if (chart.IsBubble && chart.LegendPosition != null)
            {
                bottomReserve = 0.173f;
            }
            if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
            {
                leftReserve += 0.057f;
            }
            if (!string.IsNullOrEmpty(chart.CategoryAxisTitle))
            {
                bottomReserve += 0.116f;
            }

            // Tick, title and legend bands stop growing when a chart spans extra rows or columns.
            var bandWidth = Math.Min(area.Width, 384.4f);
            var bandHeight = Math.Min(area.Height, 195.4f);
            var rightBandWidth = dateAxis ? area.Width : Math.Min(area.Width, 400f);
            return new SKRect(area.Left + bandWidth * leftReserve,
                area.Top + bandHeight * 0.187f,
                area.Right - rightBandWidth * rightReserve,
                area.Bottom - bandHeight * bottomReserve);
        }
    }
}
