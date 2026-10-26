using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class DoughnutChartRenderer
    {
        public static SKRect Draw(SKCanvas canvas, SKRect area, ParsedChart chart)
        {
            var size = Math.Min(area.Width, area.Height) * 0.635f;
            var rect = new SKRect(area.MidX - size * 0.5f, area.Top + area.Height * 0.19f,
                area.MidX + size * 0.5f, area.Top + area.Height * 0.19f + size);
            if (canvas == null || chart == null || chart.Series.Count == 0)
            {
                return rect;
            }

            var outerRadius = size * 0.5f;
            var holeRadius = outerRadius * (float)(chart.DoughnutHoleSizePercent / 100d);
            var ringWidth = (outerRadius - holeRadius) / chart.Series.Count;
            var centerX = rect.MidX;
            var centerY = rect.MidY;

            for (var ringIndex = 0; ringIndex < chart.Series.Count; ringIndex++)
            {
                var seriesIndex = chart.Series.Count - ringIndex - 1;
                var values = chart.Series[seriesIndex].Values;
                double sum = 0d;
                for (var i = 0; i < values.Count; i++)
                {
                    if (values[i].HasValue)
                    {
                        sum += Math.Abs(values[i].Value);
                    }
                }
                if (sum <= 0d)
                {
                    continue;
                }

                var ringOuter = outerRadius - ringIndex * ringWidth;
                var ringInner = ringOuter - ringWidth;
                var outerRect = SKRect.Create(centerX - ringOuter, centerY - ringOuter,
                    ringOuter * 2f, ringOuter * 2f);
                var innerRect = SKRect.Create(centerX - ringInner, centerY - ringInner,
                    ringInner * 2f, ringInner * 2f);
                var start = (float)(-90d + chart.DoughnutFirstSliceAngleDeg);
                for (var i = 0; i < values.Count; i++)
                {
                    if (!values[i].HasValue)
                    {
                        continue;
                    }

                    var sweep = (float)(Math.Abs(values[i].Value) / sum * 360d);
                    var startRadians = start * (float)Math.PI / 180f;
                    var endRadians = (start + sweep) * (float)Math.PI / 180f;
                    var color = i < chart.PiePointColors.Count ? chart.PiePointColors[i] : chart.Series[seriesIndex].Color;
                    using (var path = new SKPath())
                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = color, IsAntialias = true })
                    {
                        path.MoveTo(centerX + ringOuter * (float)Math.Cos(startRadians),
                            centerY + ringOuter * (float)Math.Sin(startRadians));
                        path.ArcTo(outerRect, start, sweep, false);
                        path.LineTo(centerX + ringInner * (float)Math.Cos(endRadians),
                            centerY + ringInner * (float)Math.Sin(endRadians));
                        path.ArcTo(innerRect, start + sweep, -sweep, false);
                        path.Close();
                        canvas.DrawPath(path, fill);
                    }

                    start += sweep;
                }
            }

            return rect;
        }
    }
}
