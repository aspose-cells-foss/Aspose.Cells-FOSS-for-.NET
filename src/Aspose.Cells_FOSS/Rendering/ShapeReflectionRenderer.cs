using System;
using System.Globalization;
using System.Xml.Linq;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapeReflectionRenderer
    {
        public static void Draw(SKCanvas canvas, SKPoint[] points, ParsedShape shape)
        {
            var reflection = shape.Reflection;
            if (reflection == null || points.Length < 2) return;
            // Vertical, below-shape reflection. Other reflection transforms are not handled here.
            if (Number(reflection, "sy", -100000d) >= 0d || Math.Abs(Number(reflection, "dir", 5400000d) - 5400000d) > 1d) return;
            var top = points[0].Y;
            var bottom = top;
            for (var i = 1; i < points.Length; i++)
            {
                top = Math.Min(top, points[i].Y);
                bottom = Math.Max(bottom, points[i].Y);
            }
            var height = bottom - top;
            if (height <= 0f) return;
            var distance = (float)(Number(reflection, "dist", 0d) / 12700d);
            var endPosition = (float)Math.Max(0.001d, Math.Min(0.999d, Number(reflection, "endPos", 100000d) / 100000d));
            var startColor = shape.LineColor.WithAlpha(Alpha(reflection, "stA", 100000d));
            var endColor = shape.LineColor.WithAlpha(Alpha(reflection, "endA", 0d));
            using (var shader = SKShader.CreateLinearGradient(new SKPoint(0f, bottom + distance), new SKPoint(0f, bottom + distance + height),
                new[] { startColor, endColor, shape.LineColor.WithAlpha(0) }, new[] { 0f, endPosition, 1f }, SKShaderTileMode.Clamp))
            using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, StrokeWidth = (float)shape.LineWidthPt, Shader = shader, IsAntialias = true })
            using (var path = new SKPath())
            {
                for (var i = 0; i < points.Length; i++)
                {
                    var point = new SKPoint(points[i].X, 2f * bottom + distance - points[i].Y);
                    if (i == 0) path.MoveTo(point);
                    else path.LineTo(point);
                }
                canvas.DrawPath(path, paint);
            }
        }

        private static byte Alpha(XElement element, string name, double fallback)
        {
            return (byte)Math.Round(Math.Max(0d, Math.Min(100000d, Number(element, name, fallback))) * 255d / 100000d);
        }

        private static double Number(XElement element, string name, double fallback)
        {
            double value;
            return double.TryParse((string)element.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                && !double.IsNaN(value) && !double.IsInfinity(value) ? value : fallback;
        }
    }
}
