using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Renders a ChartEx region map from worksheet values. Excel's opaque Bing geo-cache is not
    /// decoded; the bundled public-domain Natural Earth outlines supply the country geometry.
    /// </summary>
    internal sealed class MapChartRenderer
    {
        private readonly RenderContext _context;

        public MapChartRenderer(RenderContext context)
        {
            _context = context;
        }

        public void Draw(SKCanvas canvas, SKRect area, MapChartData chart, PageLayout page = null)
        {
            if (chart == null || area.Width <= 0f || area.Height <= 0f)
            {
                return;
            }

            canvas.Save();
            canvas.ClipRect(area);
            using (var border = new SKPaint { Color = new SKColor(0xD9, 0xD9, 0xD9), Style = SKPaintStyle.Stroke, StrokeWidth = 0.55f })
            {
                canvas.DrawRect(area, border);
            }

            DrawCenteredText(canvas, page, area, chart.Title, area.MidX, area.Top + area.Height * 0.078f,
                14f, new SKColor(0x59, 0x59, 0x59));

            float legendNameWidth;
            using (var legendFont = CreateTextPaint(8f, new SKColor(0x59, 0x59, 0x59)))
            {
                legendNameWidth = legendFont.MeasureText(chart.SeriesName ?? string.Empty);
            }

            var compactLegend = Math.Max(0f, Math.Min(1f, (70f - legendNameWidth) / 53f));
            var map = new SKRect(area.Left + area.Width * (0.11f + 0.045f * compactLegend),
                area.Top + area.Height * (0.296f - 0.066f * compactLegend),
                area.Left + area.Width * (0.678f + 0.197f * compactLegend),
                area.Top + area.Height * (0.81f + 0.055f * compactLegend));
            DrawCountries(canvas, map, chart);
            DrawLegend(canvas, page, area, chart, compactLegend);
            DrawCenteredText(canvas, page, area, "Map data: Natural Earth", area.MidX,
                area.Bottom - area.Height * 0.043f, 5.5f, new SKColor(0x66, 0x66, 0x66));
            canvas.Restore();
        }

        private void DrawCountries(SKCanvas canvas, SKRect map, MapChartData chart)
        {
            var assembly = typeof(MapChartRenderer).GetTypeInfo().Assembly;
            using (var stream = assembly.GetManifestResourceStream("Aspose.Cells.FOSS.Rendering.MapWorld110m.txt"))
            {
                if (stream == null)
                {
                    return;
                }

                using (var reader = new StreamReader(stream))
                using (var fill = new SKPaint { IsAntialias = true, Style = SKPaintStyle.Fill })
                using (var outline = new SKPaint { IsAntialias = true, Style = SKPaintStyle.Stroke,
                    StrokeWidth = 0.18f, Color = SKColors.White })
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.Length == 0 || line[0] == '#')
                        {
                            continue;
                        }

                        var separator = line.IndexOf('|');
                        if (separator < 0)
                        {
                            continue;
                        }

                        double value;
                        var country = line.Substring(0, separator);
                        var matched = chart.Values.TryGetValue(country, out value);
                        if (!matched && string.Equals(country, "United States of America", StringComparison.OrdinalIgnoreCase))
                        {
                            matched = chart.Values.TryGetValue("United States", out value);
                        }

                        fill.Color = matched ? ValueColor(value, chart)
                            : new SKColor(0xDF, 0xDF, 0xDF);
                        using (var path = BuildPath(line.Substring(separator + 1), map))
                        {
                            if (path == null)
                            {
                                continue;
                            }

                            canvas.DrawPath(path, fill);
                            canvas.DrawPath(path, outline);
                        }
                    }
                }
            }
        }

        private static SKPath BuildPath(string coordinates, SKRect map)
        {
            var points = coordinates.Split(' ');
            if (points.Length < 3)
            {
                return null;
            }

            var path = new SKPath();
            var count = 0;
            for (var i = 0; i < points.Length; i++)
            {
                var comma = points[i].IndexOf(',');
                if (comma <= 0)
                {
                    continue;
                }

                float longitude;
                float latitude;
                if (!float.TryParse(points[i].Substring(0, comma), NumberStyles.Float, CultureInfo.InvariantCulture, out longitude)
                    || !float.TryParse(points[i].Substring(comma + 1), NumberStyles.Float, CultureInfo.InvariantCulture, out latitude))
                {
                    continue;
                }

                var x = map.Left + (longitude + 180f) / 360f * map.Width;
                var y = map.Top + (90f - latitude) / 180f * map.Height;
                if (count == 0)
                {
                    path.MoveTo(x, y);
                }
                else
                {
                    path.LineTo(x, y);
                }

                count++;
            }

            if (count < 3)
            {
                path.Dispose();
                return null;
            }

            path.Close();
            return path;
        }

        private static SKColor ValueColor(double value, MapChartData chart)
        {
            var fraction = chart.Maximum > chart.Minimum
                ? (value - chart.Minimum) / (chart.Maximum - chart.Minimum) : 0.5d;
            fraction = Math.Max(0d, Math.Min(1d, fraction));
            if (chart.HasCustomColors)
            {
                if (chart.HasMiddleColor)
                {
                    return fraction <= 0.5d
                        ? Blend(chart.MinimumColor, chart.MiddleColor, fraction * 2d)
                        : Blend(chart.MiddleColor, chart.MaximumColor, (fraction - 0.5d) * 2d);
                }

                return Blend(chart.MinimumColor, chart.MaximumColor, fraction);
            }

            return new SKColor((byte)(209d - 175d * fraction),
                (byte)(218d - 137d * fraction),
                (byte)(237d - 80d * fraction));
        }

        private static SKColor Blend(SKColor first, SKColor second, double fraction)
        {
            return new SKColor(
                (byte)Math.Round(first.Red + (second.Red - first.Red) * fraction),
                (byte)Math.Round(first.Green + (second.Green - first.Green) * fraction),
                (byte)Math.Round(first.Blue + (second.Blue - first.Blue) * fraction));
        }

        private void DrawLegend(SKCanvas canvas, PageLayout page, SKRect area, MapChartData chart, float compactLegend)
        {
            var x = area.Left + area.Width * (0.737f + 0.173f * compactLegend);
            var top = area.Top + area.Height * 0.211f;
            DrawLeftText(canvas, page, area, chart.SeriesName, x - 3f, top, 8f);

            var bar = new SKRect(x, top + area.Height * 0.047f,
                x + area.Width * 0.022f, top + area.Height * 0.32f);
            using (var paint = new SKPaint())
            {
                paint.Shader = SKShader.CreateLinearGradient(new SKPoint(bar.MidX, bar.Top),
                    new SKPoint(bar.MidX, bar.Bottom),
                    chart.HasMiddleColor
                        ? new[] { ValueColor(chart.Maximum, chart), ValueColor((chart.Minimum + chart.Maximum) * 0.5d, chart), ValueColor(chart.Minimum, chart) }
                        : new[] { ValueColor(chart.Maximum, chart), ValueColor(chart.Minimum, chart) },
                    chart.HasMiddleColor ? new[] { 0f, 0.5f, 1f } : null,
                    SKShaderTileMode.Clamp);
                canvas.DrawRect(bar, paint);
            }

            DrawLeftText(canvas, page, area, string.IsNullOrEmpty(chart.MaximumLabel)
                ? chart.Maximum.ToString("0.000%", CultureInfo.InvariantCulture) : chart.MaximumLabel,
                bar.Right + 3f, bar.Top + 3f, 8f);
            DrawLeftText(canvas, page, area, string.IsNullOrEmpty(chart.MinimumLabel)
                ? chart.Minimum.ToString("0.000%", CultureInfo.InvariantCulture) : chart.MinimumLabel,
                bar.Right + 3f, bar.Bottom + 3f, 8f);
        }

        private void DrawCenteredText(SKCanvas canvas, PageLayout page, SKRect clip, string text, float centerX, float baselineY,
            float size, SKColor color)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            using (var paint = CreateTextPaint(size, color))
            {
                DrawText(canvas, page, clip, text, centerX - paint.MeasureText(text) * 0.5f, baselineY, paint);
            }
        }

        private void DrawLeftText(SKCanvas canvas, PageLayout page, SKRect clip, string text, float x, float baselineY, float size)
        {
            using (var paint = CreateTextPaint(size, new SKColor(0x59, 0x59, 0x59)))
            {
                DrawText(canvas, page, clip, text, x, baselineY, paint);
            }
        }

        private void DrawText(SKCanvas canvas, PageLayout page, SKRect clip, string text,
            float x, float baselineY, SKPaint paint)
        {
            var session = _context.PdfTextSession;
            if (session != null && page != null)
            {
                session.RecordChartRun(page, clip, text, paint.Typeface, paint.Color,
                    paint.TextSize, x, baselineY);
                return;
            }

            canvas.DrawText(text, x, baselineY, paint);
        }

        private SKPaint CreateTextPaint(float size, SKColor color)
        {
            var font = new FontValue { Name = "Calibri", Size = size };
            return new SKPaint { IsAntialias = true, Typeface = _context.Fonts.Resolve(font),
                TextSize = size, Color = color, Style = SKPaintStyle.Fill };
        }
    }
}
