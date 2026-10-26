using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartValueAxisLineRenderer
    {
        public static void Draw(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, double major)
        {
            if (!chart.HasValueAxisLine || max <= min || major <= 0d)
            {
                return;
            }

            var axisX = chart.ValueAxisOnRight ? plot.Right : plot.Left;
            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.ValueAxisLineColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.ValueAxisLineWidthPt),
                IsAntialias = false
            })
            {
                canvas.DrawLine(axisX, plot.Bottom, axisX, plot.Top, paint);

                var tickMark = chart.ValueAxisMajorTickMark;
                if (tickMark != "cross" && tickMark != "in" && tickMark != "out")
                {
                    return;
                }

                var inward = chart.ValueAxisOnRight ? -2.8f : 2.8f;
                var start = tickMark == "cross" || tickMark == "out" ? -inward : 0f;
                var end = tickMark == "cross" || tickMark == "in" ? inward : 0f;
                var steps = (int)Math.Round((max - min) / major);
                for (var i = 0; i <= steps; i++)
                {
                    var value = min + i * major;
                    var y = plot.Bottom - (float)((value - min) / (max - min)) * plot.Height;
                    canvas.DrawLine(axisX + start, y, axisX + end, y, paint);
                }
            }
        }
    }
}
