using System;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Draws classic SpreadsheetML pattern fills using foreground/background colors and a small
    /// tiled geometry, instead of collapsing every non-solid fill to a solid foreground color.
    /// </summary>
    internal static class PatternFillRenderer
    {
        public static bool Draw(SKCanvas canvas, SKRect rect, StyleValue style, RenderColor colors)
        {
            if (canvas == null || colors == null || style == null || rect.Width <= 0f || rect.Height <= 0f)
            {
                return false;
            }

            if (style.Pattern == FillPatternKind.None)
            {
                return false;
            }

            if (style.Pattern == FillPatternKind.Solid)
            {
                var solid = colors.Resolve(style.ForegroundColor, SKColors.Transparent);
                if (solid.Alpha == 0)
                {
                    solid = colors.Resolve(style.BackgroundColor, SKColors.Transparent);
                }

                if (solid.Alpha == 0)
                {
                    return false;
                }

                using (var paint = new SKPaint())
                {
                    paint.Style = SKPaintStyle.Fill;
                    paint.IsAntialias = false;
                    paint.Color = solid;
                    canvas.DrawRect(rect, paint);
                }

                return true;
            }

            var background = colors.Resolve(style.BackgroundColor, SKColors.White);
            var foreground = colors.Resolve(style.ForegroundColor, SKColors.Black);

            using (var backgroundPaint = new SKPaint())
            {
                backgroundPaint.Style = SKPaintStyle.Fill;
                backgroundPaint.IsAntialias = false;
                backgroundPaint.Color = background;
                canvas.DrawRect(rect, backgroundPaint);
            }

            if (foreground.Alpha == 0)
            {
                return background.Alpha != 0;
            }

            canvas.Save();
            canvas.ClipRect(rect);

            switch (style.Pattern)
            {
                case FillPatternKind.MediumGray:
                    DrawDensityPattern(canvas, rect, foreground, background, 0.40f, 2.5f, 2.5f, 0.15f, 50, true);
                    break;
                case FillPatternKind.DarkGray:
                    DrawDensityHolePattern(canvas, rect, foreground, background, 0.64f, 2.35f, 2.35f, 0.10f, 60, true);
                    break;
                case FillPatternKind.LightGray:
                    DrawDensityPattern(canvas, rect, foreground, background, 0.145f, 2.9f, 2.9f, 0.10f, 36, true);
                    break;
                case FillPatternKind.Gray125:
                    DrawDensityPattern(canvas, rect, foreground, background, 0.065f, 3.45f, 3.45f, 0.10f, 38, true);
                    break;
                case FillPatternKind.Gray0625:
                    DrawDensityPattern(canvas, rect, foreground, background, 0.028f, 4.6f, 4.6f, 0.09f, 36, true);
                    break;
                case FillPatternKind.DarkHorizontal:
                    DrawStripedPattern(canvas, rect, foreground, background, 0.00f, 1.0f, 0.34f, true, 255);
                    break;
                case FillPatternKind.DarkVertical:
                    DrawStripedPattern(canvas, rect, foreground, background, 0.16f, 1.56f, 0.66f, false, 244);
                    break;
                case FillPatternKind.DarkDown:
                    DrawDiagonalPattern(canvas, rect, foreground, background, 2.3f, 0.34f, true, 154);
                    break;
                case FillPatternKind.DarkUp:
                    DrawDiagonalPattern(canvas, rect, foreground, background, 2.3f, 0.34f, false, 154);
                    break;
                case FillPatternKind.DarkGrid:
                    DrawGridPattern(canvas, rect, foreground, background, 2.3f, 0.32f, 146);
                    break;
                case FillPatternKind.DarkTrellis:
                    DrawTrellisPattern(canvas, rect, foreground, background, 2.5f, 0.32f, 146);
                    break;
                case FillPatternKind.LightHorizontal:
                    DrawStripedPattern(canvas, rect, foreground, background, 0.045f, 3.0f, 0.24f, true, 136);
                    break;
                case FillPatternKind.LightVertical:
                    DrawStripedPattern(canvas, rect, foreground, background, 0.045f, 3.0f, 0.24f, false, 136);
                    break;
                case FillPatternKind.LightDown:
                    DrawDiagonalPattern(canvas, rect, foreground, background, 3.3f, 0.17f, true, 102);
                    break;
                case FillPatternKind.LightUp:
                    DrawDiagonalPattern(canvas, rect, foreground, background, 3.3f, 0.17f, false, 102);
                    break;
                case FillPatternKind.LightGrid:
                    DrawGridPattern(canvas, rect, foreground, background, 3.2f, 0.17f, 96);
                    break;
                case FillPatternKind.LightTrellis:
                    DrawTrellisPattern(canvas, rect, foreground, background, 1.05f, 0.16f, 136);
                    break;
            }

            canvas.Restore();
            return true;
        }

        private static void DrawDensityPattern(SKCanvas canvas, SKRect rect, SKColor foreground, SKColor background, float coverage, float stepX, float stepY, float dotSize, byte alpha, bool staggered)
        {
            FillRect(canvas, rect, Blend(background, foreground, coverage));
            DrawDots(canvas, rect, WithAlpha(foreground, alpha), stepX, stepY, dotSize, staggered);
        }

        private static void DrawDensityHolePattern(SKCanvas canvas, SKRect rect, SKColor foreground, SKColor background, float coverage, float stepX, float stepY, float holeSize, byte alpha, bool staggered)
        {
            FillRect(canvas, rect, Blend(background, foreground, coverage));
            DrawDots(canvas, rect, WithAlpha(background, alpha), stepX, stepY, holeSize, staggered);
        }

        private static void DrawStripedPattern(SKCanvas canvas, SKRect rect, SKColor foreground, SKColor background, float coverage, float step, float strokeWidth, bool horizontal, byte alpha)
        {
            FillRect(canvas, rect, Blend(background, foreground, coverage));
            var accent = WithAlpha(foreground, alpha);
            if (horizontal)
            {
                DrawHorizontalLines(canvas, rect, accent, step, strokeWidth);
            }
            else
            {
                DrawVerticalLines(canvas, rect, accent, step, strokeWidth);
            }
        }

        private static void DrawDiagonalPattern(SKCanvas canvas, SKRect rect, SKColor foreground, SKColor background, float step, float strokeWidth, bool down, byte alpha)
        {
            FillRect(canvas, rect, background);
            DrawDiagonalLines(canvas, rect, WithAlpha(foreground, alpha), step, strokeWidth, down);
        }

        private static void DrawGridPattern(SKCanvas canvas, SKRect rect, SKColor foreground, SKColor background, float step, float strokeWidth, byte alpha)
        {
            FillRect(canvas, rect, background);
            var accent = WithAlpha(foreground, alpha);
            DrawHorizontalLines(canvas, rect, accent, step, strokeWidth);
            DrawVerticalLines(canvas, rect, accent, step, strokeWidth);
        }

        private static void DrawTrellisPattern(SKCanvas canvas, SKRect rect, SKColor foreground, SKColor background, float step, float strokeWidth, byte alpha)
        {
            FillRect(canvas, rect, background);
            var accent = WithAlpha(foreground, alpha);
            DrawDiagonalLines(canvas, rect, accent, step, strokeWidth, true);
            DrawDiagonalLines(canvas, rect, accent, step, strokeWidth, false);
        }

        private static void FillRect(SKCanvas canvas, SKRect rect, SKColor color)
        {
            using (var paint = new SKPaint())
            {
                paint.Style = SKPaintStyle.Fill;
                paint.IsAntialias = false;
                paint.Color = color;
                canvas.DrawRect(rect, paint);
            }
        }

        private static void DrawDots(SKCanvas canvas, SKRect rect, SKColor color, float stepX, float stepY, float dotSize, bool staggered)
        {
            using (var paint = new SKPaint())
            {
                paint.Style = SKPaintStyle.Fill;
                paint.IsAntialias = false;
                paint.Color = color;

                var rowIndex = 0;
                for (var top = rect.Top; top < rect.Bottom; top += stepY)
                {
                    var offset = staggered && rowIndex % 2 == 1 ? stepX * 0.5f : 0f;
                    for (var left = rect.Left + offset; left < rect.Right; left += stepX)
                    {
                        canvas.DrawRect(left, top, dotSize, dotSize, paint);
                    }

                    rowIndex++;
                }
            }
        }

        private static void DrawHorizontalLines(SKCanvas canvas, SKRect rect, SKColor color, float step, float strokeWidth)
        {
            using (var paint = new SKPaint())
            {
                paint.Style = SKPaintStyle.Stroke;
                paint.IsAntialias = false;
                paint.Color = color;
                paint.StrokeWidth = strokeWidth;

                for (var y = rect.Top + strokeWidth * 0.5f; y < rect.Bottom; y += step)
                {
                    canvas.DrawLine(rect.Left, y, rect.Right, y, paint);
                }
            }
        }

        private static void DrawVerticalLines(SKCanvas canvas, SKRect rect, SKColor color, float step, float strokeWidth)
        {
            using (var paint = new SKPaint())
            {
                paint.Style = SKPaintStyle.Stroke;
                paint.IsAntialias = false;
                paint.Color = color;
                paint.StrokeWidth = strokeWidth;

                for (var x = rect.Left + strokeWidth * 0.5f; x < rect.Right; x += step)
                {
                    canvas.DrawLine(x, rect.Top, x, rect.Bottom, paint);
                }
            }
        }

        private static void DrawDiagonalLines(SKCanvas canvas, SKRect rect, SKColor color, float step, float strokeWidth, bool down)
        {
            using (var paint = new SKPaint())
            {
                paint.Style = SKPaintStyle.Stroke;
                paint.IsAntialias = false;
                paint.Color = color;
                paint.StrokeWidth = strokeWidth;

                var width = rect.Width;
                var height = rect.Height;
                for (var offset = -height; offset <= width; offset += step)
                {
                    if (down)
                    {
                        canvas.DrawLine(rect.Left + offset, rect.Top, rect.Left + offset + height, rect.Bottom, paint);
                    }
                    else
                    {
                        canvas.DrawLine(rect.Left + offset, rect.Bottom, rect.Left + offset + height, rect.Top, paint);
                    }
                }
            }
        }

        private static SKColor Blend(SKColor background, SKColor foreground, float foregroundFraction)
        {
            var clamped = Math.Max(0f, Math.Min(1f, foregroundFraction));
            var inverse = 1f - clamped;
            return new SKColor(
                (byte)Math.Round(background.Red * inverse + foreground.Red * clamped),
                (byte)Math.Round(background.Green * inverse + foreground.Green * clamped),
                (byte)Math.Round(background.Blue * inverse + foreground.Blue * clamped),
                (byte)Math.Round(background.Alpha * inverse + foreground.Alpha * clamped));
        }

        private static SKColor WithAlpha(SKColor color, byte alpha)
        {
            return new SKColor(color.Red, color.Green, color.Blue, alpha);
        }
    }
}
