using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class BubbleChartMath
    {
        public static void ResolveXRange(ParsedChart chart, double dataMinimum, double dataMaximum,
            out double minimum, out double maximum, out double major)
        {
            if (dataMinimum >= 0d && !ChartXmlParser.IsDateFormat(chart.ScatterXFormatCode)
                && !chart.ScatterXMin.HasValue && !chart.ScatterXMax.HasValue)
            {
                major = chart.ScatterXMajorUnit ?? NiceNumber(dataMaximum / 8d);
                if (major <= 0d)
                {
                    major = 1d;
                }
                minimum = 0d;
                maximum = Math.Max(Math.Ceiling(dataMaximum / major), 8d) * major;
                return;
            }

            ResolveRange(dataMinimum, dataMaximum, chart.ScatterXMin, chart.ScatterXMax,
                chart.ScatterXMajorUnit, out minimum, out maximum, out major);
        }

        public static void ResolveYRange(ParsedChart chart, double dataMinimum, double dataMaximum,
            out double minimum, out double maximum, out double major)
        {
            ResolveRange(dataMinimum, dataMaximum, chart.AxisMin, chart.AxisMax,
                chart.MajorUnit, out minimum, out maximum, out major);
        }

        public static void DrawSeries(SKCanvas canvas, SKRect plot, ParsedChart chart,
            double minimumY, double maximumY)
        {
            double minimumX;
            double maximumX;
            double majorX;
            ScatterChartAxisMath.ResolveXRange(chart, out minimumX, out maximumX, out majorX);
            if (maximumX <= minimumX || maximumY <= minimumY)
            {
                return;
            }

            var maximumSize = 0d;
            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
            {
                var series = chart.Series[seriesIndex];
                for (var pointIndex = 0; pointIndex < series.BubbleSizes.Count; pointIndex++)
                {
                    if (series.BubbleSizes[pointIndex].HasValue)
                    {
                        var size = series.BubbleSizes[pointIndex].Value;
                        if (size > 0d || chart.ShowNegativeBubbles)
                        {
                            maximumSize = Math.Max(maximumSize, Math.Abs(size));
                        }
                    }
                }
            }
            if (maximumSize <= 0d)
            {
                return;
            }

            var maximumRadius = plot.Height * 0.1264f * (float)(chart.BubbleScalePercent / 100d);
            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
            {
                var series = chart.Series[seriesIndex];
                var count = Math.Min(Math.Min(series.XValues.Count, series.Values.Count), series.BubbleSizes.Count);
                using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = series.Color, IsAntialias = true })
                using (var texture = series.Bubble3D ? BubbleSphereRenderer.CreateTexture(series.Color) : null)
                {
                    for (var pointIndex = 0; pointIndex < count; pointIndex++)
                    {
                        if (!series.XValues[pointIndex].HasValue || !series.Values[pointIndex].HasValue
                            || !series.BubbleSizes[pointIndex].HasValue)
                        {
                            continue;
                        }

                        var xValue = series.XValues[pointIndex].Value;
                        var yValue = series.Values[pointIndex].Value;
                        var size = series.BubbleSizes[pointIndex].Value;
                        if (xValue < minimumX || xValue > maximumX || yValue < minimumY || yValue > maximumY
                            || size == 0d || size < 0d && !chart.ShowNegativeBubbles)
                        {
                            continue;
                        }

                        var x = ScatterChartAxisMath.ToScreen(xValue, minimumX, maximumX, plot.Left, plot.Right);
                        var y = plot.Bottom - (float)((yValue - minimumY) / (maximumY - minimumY)) * plot.Height;
                        var radius = maximumRadius * (float)Math.Sqrt(Math.Abs(size) / maximumSize);
                        if (series.Bubble3D)
                        {
                            BubbleSphereRenderer.DrawImage(canvas, new SKPoint(x, y), radius, texture);
                        }
                        else
                        {
                            canvas.DrawCircle(x, y, radius, fill);
                        }
                    }
                }
            }
        }

        private static void ResolveRange(double dataMinimum, double dataMaximum,
            double? explicitMinimum, double? explicitMaximum, double? explicitMajor,
            out double minimum, out double maximum, out double major)
        {
            var span = Math.Max(0d, dataMaximum - dataMinimum);
            major = explicitMajor ?? NiceNumber(span / 6d);
            if (major <= 0d)
            {
                major = 1d;
            }

            var padding = span > 0d ? span * 0.25d : major;
            minimum = explicitMinimum ?? Math.Floor((dataMinimum - padding) / major) * major;
            maximum = explicitMaximum ?? Math.Ceiling((dataMaximum + padding) / major) * major;
            if (maximum <= minimum)
            {
                maximum = minimum + major;
            }
        }

        private static double NiceNumber(double value)
        {
            if (value <= 0d)
            {
                return 1d;
            }

            var exponent = Math.Floor(Math.Log10(value));
            var fraction = value / Math.Pow(10d, exponent);
            double nice;
            if (fraction < 1.5d) nice = 1d;
            else if (fraction < 3d) nice = 2d;
            else if (fraction < 7d) nice = 5d;
            else nice = 10d;
            return nice * Math.Pow(10d, exponent);
        }
    }
}
