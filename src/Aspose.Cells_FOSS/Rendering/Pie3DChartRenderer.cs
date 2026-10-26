using System;
using System.Collections.Generic;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Draws the projected top and visible front rim of a three-dimensional pie.
    /// </summary>
    internal static class Pie3DChartRenderer
    {
        public static SKRect Draw(SKCanvas canvas, SKRect area, ParsedChart chart,
            List<double?> values, List<SKColor> colors, double sum)
        {
            var perspective = chart.RightAngleAxes ? 0f
                : (float)(Math.Sin(chart.PerspectiveDeg * Math.PI / 180d) * 0.68d);
            perspective = Math.Max(0f, Math.Min(0.6f, perspective));
            var flatView = 1f - Math.Min(1f, perspective / 0.34f);
            var projectedRadiusX = area.Width * (0.45f - 0.03f * flatView);
            var radiusX = projectedRadiusX * (float)Math.Sqrt(1f - perspective * perspective);
            var tilt = (float)Math.Sin(Math.Max(0d, Math.Min(70d, chart.RotationXDeg)) * Math.PI / 180d);
            var radiusY = area.Height * (0.15f + 0.065f * flatView) * Math.Max(0.2f,
                tilt / (float)Math.Sin(15d * Math.PI / 180d));
            var centerY = area.Top + area.Height * (0.303f + 0.107f * flatView);
            var depth = area.Height * 0.193f * (float)Math.Max(0.25d,
                Math.Min(2d, chart.DepthPercent / 100d));
            var center = new SKPoint(area.MidX, centerY);
            var ellipse = new SKRect(center.X - projectedRadiusX,
                center.Y - radiusY / (1f + perspective),
                center.X + projectedRadiusX,
                center.Y + radiusY / (1f - perspective));

            var start = -90f + (float)chart.RotationYDeg;
            var sliceStarts = new float[values.Count];
            for (var i = 0; i < values.Count; i++)
            {
                sliceStarts[i] = start;
                if (!values[i].HasValue)
                {
                    continue;
                }

                var sweep = (float)(Math.Abs(values[i].Value) / sum * 360d);
                DrawVisibleRim(canvas, center, radiusX, radiusY, perspective,
                    depth, start, sweep, colors[i]);
                start += sweep;
            }

            for (var i = values.Count - 1; i >= 0; i--)
            {
                if (!values[i].HasValue)
                {
                    continue;
                }

                var sweep = (float)(Math.Abs(values[i].Value) / sum * 360d);
                using (var path = new SKPath())
                using (var fill = new SKPaint { Style = SKPaintStyle.Fill,
                    Color = Darken(colors[i], 0.92f), IsAntialias = true })
                {
                    path.MoveTo(center);
                    AddProjectedArc(path, center, radiusX, radiusY, perspective,
                        depth, sliceStarts[i], sliceStarts[i] + sweep, false, false);
                    path.Close();
                    canvas.DrawPath(path, fill);
                }
            }

            return ellipse;
        }

        private static void DrawVisibleRim(SKCanvas canvas, SKPoint center,
            float radiusX, float radiusY, float perspective, float depth,
            float start, float sweep, SKColor color)
        {
            var silhouetteAngle = (float)(Math.Asin(perspective) * 180d / Math.PI);
            var visibleStart = Math.Max(silhouetteAngle, start);
            var visibleEnd = Math.Min(180f - silhouetteAngle, start + sweep);
            if (visibleEnd <= visibleStart)
            {
                return;
            }

            using (var path = new SKPath())
            {
                AddProjectedArc(path, center, radiusX, radiusY, perspective,
                    depth, visibleStart, visibleEnd, false, true);
                AddProjectedArc(path, center, radiusX, radiusY, perspective,
                    depth, visibleEnd, visibleStart, true, false);
                path.Close();

                var leftShade = Darken(color, 0.95f);
                var centerShade = Darken(color, 0.70f);
                var rightShade = Darken(color, 0.35f);
                using (var shader = SKShader.CreateLinearGradient(
                    new SKPoint(center.X - radiusX, 0f),
                    new SKPoint(center.X + radiusX, 0f),
                    new[] { leftShade, centerShade, rightShade },
                    new[] { 0f, 0.5f, 1f }, SKShaderTileMode.Clamp))
                using (var fill = new SKPaint { Style = SKPaintStyle.Fill,
                    Shader = shader, IsAntialias = true })
                {
                    canvas.DrawPath(path, fill);
                }
            }
        }

        private static void AddProjectedArc(SKPath path, SKPoint center,
            float radiusX, float radiusY, float perspective, float depth,
            float start, float end, bool lower, bool moveToStart)
        {
            var segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(end - start) / 2f));
            for (var i = 0; i <= segments; i++)
            {
                var angle = start + (end - start) * i / segments;
                var radians = angle * (float)Math.PI / 180f;
                var sine = (float)Math.Sin(radians);
                var divisor = 1f - perspective * sine;
                var x = center.X + radiusX * (float)Math.Cos(radians) / divisor;
                var y = center.Y + (radiusY * sine + (lower ? depth : 0f)) / divisor;
                if (lower)
                {
                    x = center.X + (x - center.X) * 0.978f;
                }

                if (i == 0 && moveToStart)
                {
                    path.MoveTo(x, y);
                }
                else
                {
                    path.LineTo(x, y);
                }
            }
        }

        private static SKColor Darken(SKColor color, float factor)
        {
            return new SKColor((byte)(color.Red * factor),
                (byte)(color.Green * factor), (byte)(color.Blue * factor), color.Alpha);
        }
    }
}
