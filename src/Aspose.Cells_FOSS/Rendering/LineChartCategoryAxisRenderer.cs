using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class LineChartCategoryAxisRenderer
    {
        public static bool DrawBoundaryGridlines(SKCanvas canvas, SKRect plot, ParsedChart chart, DateSystem dates)
        {
            if (!chart.HasCategoryGridlines || chart.Kind != ChartKind.Line || chart.IsLine3D
                || !chart.HasDateCategoryAxis || !chart.CrossBetween
                || !string.Equals(chart.CategoryBaseTimeUnit, "months", StringComparison.Ordinal))
            {
                return false;
            }
            var positions = TickPositions(plot, chart, dates);
            using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, Color = chart.CategoryGridlineColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.CategoryGridlineWidthPt), IsAntialias = false })
            {
                for (var i = 0; i < positions.Count; i++)
                {
                    canvas.DrawLine(positions[i], plot.Top, positions[i], plot.Bottom, paint);
                }
            }
            return true;
        }

        public static void Draw(SKCanvas canvas, SKRect plot, ParsedChart chart, DateSystem dates)
        {
            if (!chart.HasCategoryAxisLine) return;
            using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, Color = chart.CategoryAxisLineColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.CategoryAxisLineWidthPt), IsAntialias = false })
            {
                canvas.DrawLine(plot.Left, plot.Bottom, plot.Right, plot.Bottom, paint);
                var inside = chart.CategoryAxisMajorTickMark == "in" || chart.CategoryAxisMajorTickMark == "cross";
                var outside = chart.CategoryAxisMajorTickMark == "out" || chart.CategoryAxisMajorTickMark == "cross";
                if (!inside && !outside) return;
                var positions = TickPositions(plot, chart, dates);
                for (var i = 0; i < positions.Count; i++)
                {
                    canvas.DrawLine(positions[i], plot.Bottom - (inside ? 2.8f : 0f),
                        positions[i], plot.Bottom + (outside ? 2.8f : 0f), paint);
                }
            }
        }

        public static List<float> TickPositions(SKRect plot, ParsedChart chart, DateSystem dates)
        {
            var positions = new List<float>();
            if (!chart.HasDateCategoryAxis)
            {
                var count = chart.Categories.Count;
                var intervals = chart.CrossBetween ? count : count - 1;
                if (intervals <= 0) return positions;
                for (var i = 0; i <= intervals; i++)
                {
                    positions.Add(plot.Left + plot.Width * i / intervals);
                }
                return positions;
            }
            if (!string.Equals(chart.CategoryBaseTimeUnit, "months", StringComparison.Ordinal)) return positions;
            double min, max, plotMin, plotMax;
            if (!ChartDateAxisMath.TryGetDateAxisRange(chart, out min, out max)
                || !ChartDateAxisMath.TryGetDateAxisPlotRange(chart, dates, out plotMin, out plotMax)) return positions;
            var current = ChartDateAxisMath.StartOfMonth(min, dates);
            var last = ChartDateAxisMath.StartOfMonth(max, dates);
            if (chart.CrossBetween) last = ChartDateAxisMath.AddMonthsSerial(last, 1, dates);
            while (current <= last + 0.1d)
            {
                var serial = current;
                if (chart.CrossBetween)
                {
                    serial = (ChartDateAxisMath.AddMonthsSerial(current, -1, dates) + current) * 0.5d;
                }
                positions.Add(ChartDateAxisMath.DateCategoryX(serial, plotMin, plotMax, plot, chart, dates));
                current = ChartDateAxisMath.AddMonthsSerial(current, 1, dates);
            }
            return positions;
        }
    }
}
