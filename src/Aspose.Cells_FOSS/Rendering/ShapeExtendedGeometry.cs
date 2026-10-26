using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapeExtendedGeometry
    {
        public static SKPath Build(ParsedShape shape, SKRect r)
        {
            if (r.Width <= 0f || r.Height <= 0f) return null;
            if (shape.Geometry == "leftArrowCallout") return LeftArrowCallout(shape, r);
            if (shape.Geometry == "mathNotEqual") return NotEqual(shape, r);
            if (shape.Geometry == "flowChartMagneticDrum")
            {
                var path = new SKPath();
                path.MoveTo(r.Left + r.Width / 6f, r.Top);
                path.LineTo(r.Right - r.Width / 6f, r.Top);
                path.ArcTo(new SKRect(r.Right - r.Width / 3f, r.Top, r.Right, r.Bottom), -90f, 180f, false);
                path.LineTo(r.Left + r.Width / 6f, r.Bottom);
                path.ArcTo(new SKRect(r.Left, r.Top, r.Left + r.Width / 3f, r.Bottom), 90f, 180f, false);
                path.Close();
                return path;
            }
            return null;
        }

        public static float Adjustment(ParsedShape shape, int index, float fallback)
        {
            var value = ShapeCurveGeometry.Adjustment(shape, index, fallback);
            return float.IsNaN(value) || float.IsInfinity(value) ? fallback : value;
        }

        private static float Clamp(float value, float maximum)
        {
            return Math.Max(0f, Math.Min(maximum, value));
        }

        private static SKPath LeftArrowCallout(ParsedShape shape, SKRect r)
        {
            var ss = Math.Min(r.Width, r.Height);
            var a2 = Clamp(Adjustment(shape, 1, 0.25f), r.Height / (2f * ss));
            var a1 = Clamp(Adjustment(shape, 0, 0.25f), a2 * 2f);
            var a3 = Clamp(Adjustment(shape, 2, 0.25f), r.Width / ss);
            var a4 = Clamp(Adjustment(shape, 3, 0.64977f), 1f - a3 * ss / r.Width);
            var x1 = r.Left + ss * a3;
            var x2 = r.Right - r.Width * a4;
            var y1 = r.MidY - ss * a2;
            var y2 = r.MidY - ss * a1 / 2f;
            var y3 = r.MidY + ss * a1 / 2f;
            var y4 = r.MidY + ss * a2;
            return Polygon(new float[] { r.Left,r.MidY, x1,y1, x1,y2, x2,y2, x2,r.Top,
                r.Right,r.Top, r.Right,r.Bottom, x2,r.Bottom, x2,y3, x1,y3, x1,y4 });
        }

        private static SKPath NotEqual(ParsedShape shape, SKRect r)
        {
            var w = r.Width;
            var h = r.Height;
            var a1 = Clamp(Adjustment(shape, 0, 0.2352f), 0.5f);
            var a3 = Clamp(Adjustment(shape, 2, 0.1176f), 1f - 2f * a1);
            // The second adjustment is an angle in 60000ths of a degree;
            // parsed adjustments retain the common /100000 scale.
            var angle = Math.Max(42f, Math.Min(66f, Adjustment(shape, 1, 66f))) / 0.6f - 90f;
            var dx = h / 2f * (float)Math.Tan(angle * Math.PI / 180d);
            var length = (float)Math.Sqrt(dx * dx + h * h / 4f);
            var thickness = h * a1;
            var bandWidth = length * thickness / (h / 2f);
            var x1 = w * (0.5f - 0.36745f);
            var x8 = w - x1;
            var y2 = h * (0.5f - a3 / 2f);
            var y3 = h - y2;
            var y1 = y2 - thickness;
            var y4 = y3 + thickness;
            var x7 = w / 2f + dx - bandWidth / 2f;
            var x6 = x7 - dx * y1 / (h / 2f);
            var x5 = x7 - dx * y2 / (h / 2f);
            var x4 = x7 - dx * y3 / (h / 2f);
            var x3 = x7 - dx * y4 / (h / 2f);
            var dx7 = thickness * h / 2f / length;
            var dy = thickness * dx / length;
            var rx = angle > 0f ? x7 + dx7 : x7 + bandWidth;
            var lx = angle > 0f ? x7 : x7 + bandWidth - dx7;
            var ry = angle > 0f ? dy : 0f;
            var ly = angle > 0f ? 0f : -dy;
            var points = new float[] { x1,y1, x6,y1, lx,ly, rx,ry, x6+bandWidth,y1,
                x8,y1, x8,y2, x5+bandWidth,y2, x4+bandWidth,y3, x8,y3, x8,y4,
                x3+bandWidth,y4, w-lx,h-ly, w-rx,h-ry, x3,y4, x1,y4, x1,y3, x4,y3, x5,y2, x1,y2 };
            for (var i = 0; i < points.Length; i += 2)
            {
                points[i] += r.Left;
                points[i + 1] += r.Top;
            }
            return Polygon(points);
        }

        public static SKPath Polygon(float[] points)
        {
            var path = new SKPath();
            path.MoveTo(points[0], points[1]);
            for (var i = 2; i < points.Length; i += 2) path.LineTo(points[i], points[i + 1]);
            path.Close();
            return path;
        }
    }
}
