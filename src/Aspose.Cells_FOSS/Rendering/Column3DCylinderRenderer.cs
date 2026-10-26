using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class Column3DCylinderRenderer
    {
        private const int ArcSteps = 16;

        public static SKPoint ResolveBottomCenterOffset(SKPoint floorMidpoint, SKPoint depth)
        {
            return new SKPoint(floorMidpoint.X - depth.X * 0.5f,
                floorMidpoint.Y - depth.Y * 0.5f);
        }

        public static float ResolveHeightExtension(float floorOffsetY, float width,
            SKPoint depth, float horizontalAxisSlope)
        {
            var widthAxisY = width * horizontalAxisSlope * 0.5f;
            var depthAxisY = depth.Y * 0.5f;
            var verticalRadius = (float)Math.Sqrt(
                widthAxisY * widthAxisY + depthAxisY * depthAxisY);
            return Math.Max(0f, floorOffsetY + depthAxisY - verticalRadius);
        }

        public static void Draw(SKCanvas canvas, SKPoint bottomCenter, float width,
            SKPoint heightOffset, SKPoint depth, float horizontalAxisSlope,
            ChartSeries series, bool drawTopFace)
        {
            if (canvas == null || series == null || width <= 0f)
            {
                return;
            }

            var widthAxis = new SKPoint(width * 0.5f, width * horizontalAxisSlope * 0.5f);
            var depthAxis = new SKPoint(depth.X * 0.5f, depth.Y * 0.5f);
            var bottom = Offset(bottomCenter, depthAxis);
            var top = Offset(bottom, heightOffset);
            var extremeAngle = ResolveHorizontalExtremeAngle(widthAxis, depthAxis);
            var left = EllipsePoint(top, widthAxis, depthAxis, extremeAngle + Math.PI);
            var right = EllipsePoint(top, widthAxis, depthAxis, extremeAngle);

            var colors = new SKColor[]
            {
                Shade(series.Color, 0.72d),
                series.Color,
                Shade(series.Color, 0.72d)
            };
            var positions = new float[] { 0f, 0.48f, 1f };
            using (var shader = SKShader.CreateLinearGradient(
                left,
                right,
                colors,
                positions,
                SKShaderTileMode.Clamp))
            using (var bodyPaint = new SKPaint
            {
                Style = SKPaintStyle.Fill,
                Shader = shader,
                IsAntialias = true
            })
            using (var capPaint = new SKPaint
            {
                Style = SKPaintStyle.Fill,
                Color = Shade(series.Color, 1.07d),
                IsAntialias = true
            })
            using (var edgePaint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = series.LineColor,
                StrokeWidth = (float)Math.Max(0.25d, series.LineWidthPt),
                IsAntialias = true
            })
            {
                using (var body = BuildBodyPath(bottom, top, widthAxis, depthAxis))
                {
                    canvas.DrawPath(body, bodyPaint);
                    if (series.HasVisibleLine)
                    {
                        canvas.DrawPath(body, edgePaint);
                    }
                }

                if (drawTopFace)
                {
                    using (var cap = BuildEllipsePath(top, widthAxis, depthAxis))
                    {
                        canvas.DrawPath(cap, capPaint);
                        if (series.HasVisibleLine)
                        {
                            canvas.DrawPath(cap, edgePaint);
                        }
                    }
                }
            }
        }

        private static SKPath BuildBodyPath(SKPoint bottom, SKPoint top,
            SKPoint widthAxis, SKPoint depthAxis)
        {
            var path = new SKPath();
            var startAngle = ResolveHorizontalExtremeAngle(widthAxis, depthAxis) + Math.PI;
            for (var index = 0; index <= ArcSteps; index++)
            {
                var angle = startAngle + Math.PI * index / ArcSteps;
                var point = EllipsePoint(top, widthAxis, depthAxis, angle);
                if (index == 0)
                {
                    path.MoveTo(point);
                }
                else
                {
                    path.LineTo(point);
                }
            }

            for (var index = ArcSteps; index >= 0; index--)
            {
                var angle = startAngle + Math.PI * index / ArcSteps;
                path.LineTo(EllipsePoint(bottom, widthAxis, depthAxis, angle));
            }
            path.Close();
            return path;
        }

        private static double ResolveHorizontalExtremeAngle(
            SKPoint widthAxis, SKPoint depthAxis)
        {
            return Math.Atan2(depthAxis.X, widthAxis.X);
        }

        private static SKPath BuildEllipsePath(SKPoint center,
            SKPoint widthAxis, SKPoint depthAxis)
        {
            var path = new SKPath();
            var steps = ArcSteps * 2;
            for (var index = 0; index <= steps; index++)
            {
                var angle = Math.PI * 2d * index / steps;
                var point = EllipsePoint(center, widthAxis, depthAxis, angle);
                if (index == 0)
                {
                    path.MoveTo(point);
                }
                else
                {
                    path.LineTo(point);
                }
            }
            path.Close();
            return path;
        }

        private static SKPoint EllipsePoint(SKPoint center,
            SKPoint widthAxis, SKPoint depthAxis, double angle)
        {
            var cosine = (float)Math.Cos(angle);
            var sine = (float)Math.Sin(angle);
            return new SKPoint(
                center.X + widthAxis.X * cosine + depthAxis.X * sine,
                center.Y + widthAxis.Y * cosine + depthAxis.Y * sine);
        }

        private static SKPoint Offset(SKPoint point, SKPoint offset)
        {
            return new SKPoint(point.X + offset.X, point.Y + offset.Y);
        }

        private static SKColor Shade(SKColor color, double factor)
        {
            var red = Math.Max(0d, Math.Min(255d, color.Red * factor));
            var green = Math.Max(0d, Math.Min(255d, color.Green * factor));
            var blue = Math.Max(0d, Math.Min(255d, color.Blue * factor));
            return new SKColor((byte)Math.Round(red), (byte)Math.Round(green),
                (byte)Math.Round(blue), color.Alpha);
        }
    }
}
