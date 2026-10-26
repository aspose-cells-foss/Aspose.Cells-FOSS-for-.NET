using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartPatternRenderer
    {
        public static bool DrawChartArea(SKCanvas canvas, SKRect rect, ParsedChart chart)
        {
            if (canvas == null || chart == null || rect.Width <= 0f || rect.Height <= 0f
                || !string.Equals(chart.ChartAreaPatternPreset, "weave", StringComparison.Ordinal))
            {
                return false;
            }

            var rows = new string[]
            {
                "...#....##..", "...#....##..", ".##.##....#.", "#.....#....#",
                "#.....#....#", "....##.#..#.", "...#....##..", "...#....##..",
                ".##.##......", "#.....#....#", "#.....#....#", ".##....#..#."
            };
            var foreground = chart.ChartAreaPatternForegroundColor;
            var background = chart.ChartAreaPatternBackgroundColor;
            using (var tile = new SKBitmap(12, 12, SKColorType.Rgba8888, SKAlphaType.Premul))
            {
                for (var y = 0; y < rows.Length; y++)
                {
                    for (var x = 0; x < rows[y].Length; x++)
                    {
                        var source = rows[y][x] == '#' ? foreground : background;
                        // Near-opaque pixels keep Skia's PDF tile lossless instead of JPEG-compressing away the blue detail.
                        tile.SetPixel(x, y, new SKColor(source.Red, source.Green, source.Blue, 254));
                    }
                }

                using (var shader = tile.ToShader(SKShaderTileMode.Repeat, SKShaderTileMode.Repeat,
                    SKMatrix.CreateScale(0.5f, 0.5f)))
                using (var paint = new SKPaint { Style = SKPaintStyle.Fill, Shader = shader, IsAntialias = false })
                {
                    canvas.DrawRect(rect, paint);
                }
            }

            return true;
        }

        public static bool DrawPath(SKCanvas canvas, SKPath path, ChartSeries series, SKColor fallbackColor)
        {
            if (canvas == null || path == null || series == null
                || !string.Equals(series.PatternPreset, "ltUpDiag", StringComparison.Ordinal))
            {
                return false;
            }

            var bounds = path.Bounds;
            var background = series.PatternBackgroundColor.Alpha > 0
                ? series.PatternBackgroundColor
                : fallbackColor;
            var foreground = series.PatternForegroundColor.Alpha > 0
                ? series.PatternForegroundColor
                : series.Color;

            using (var backgroundPaint = new SKPaint { Style = SKPaintStyle.Fill, Color = background, IsAntialias = true })
            {
                canvas.DrawPath(path, backgroundPaint);
            }

            canvas.Save();
            canvas.ClipPath(path, SKClipOperation.Intersect, true);
            DrawLightUpDiagonal(canvas, bounds, foreground);
            canvas.Restore();
            return true;
        }

        public static bool DrawRect(SKCanvas canvas, SKRect rect, ChartSeries series, SKColor fallbackColor)
        {
            if (canvas == null || series == null || rect.Width <= 0f || rect.Height <= 0f)
            {
                return false;
            }

            using (var path = new SKPath())
            {
                path.AddRect(rect);
                return DrawPath(canvas, path, series, fallbackColor);
            }
        }

        private static void DrawLightUpDiagonal(SKCanvas canvas, SKRect bounds, SKColor color)
        {
            var lineColor = new SKColor(color.Red, color.Green, color.Blue, 150);
            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = lineColor,
                StrokeWidth = 0.35f,
                IsAntialias = true
            })
            {
                var height = Math.Max(1f, bounds.Height);
                for (var x = bounds.Left - height; x <= bounds.Right; x += 2.2f)
                {
                    canvas.DrawLine(x, bounds.Bottom, x + height, bounds.Top, paint);
                }
            }
        }
    }
}
