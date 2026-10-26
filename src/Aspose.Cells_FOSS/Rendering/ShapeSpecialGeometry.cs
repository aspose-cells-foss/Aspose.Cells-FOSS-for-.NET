using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapeSpecialGeometry
    {
        public static SKPath Build(ParsedShape shape, SKRect rect)
        {
            if (shape.CustomGeometry != null)
            {
                return ShapeCustomGeometry.Build(shape.CustomGeometry, rect);
            }

            if (shape.Geometry == "irregularSeal1")
            {
                // DrawingML irregularSeal1 vertices in its 21600 x 21600 coordinate space.
                var vertices = new float[] { 10800,5800, 14522,0, 14155,5325, 18380,4457,
                    16702,7315, 21097,8137, 17607,10475, 21600,13290, 16837,12942,
                    18145,18095, 14020,14457, 13247,19737, 10532,14935, 8485,21600,
                    7715,15627, 4762,17617, 5667,13937, 135,14587, 3722,11775,
                    0,8615, 4627,7617, 370,2295, 7312,6320, 8352,2295 };
                var path = new SKPath();
                for (var i = 0; i < vertices.Length; i += 2)
                {
                    var x = rect.Left + rect.Width * vertices[i] / 21600f;
                    var y = rect.Top + rect.Height * vertices[i + 1] / 21600f;
                    if (i == 0) path.MoveTo(x, y);
                    else path.LineTo(x, y);
                }
                path.Close();
                return path;
            }

            if (shape.Geometry == "wedgeRectCallout")
            {
                return BuildCallout(shape.Adjusts, rect);
            }

            if (shape.Geometry == "snip1Rect" || shape.Geometry == "snip2SameRect")
            {
                var top = Math.Min(rect.Width, rect.Height) * Math.Max(0f, Math.Min(0.5f, ShapeCurveGeometry.Adjustment(shape, 0, 0.16667f)));
                var both = shape.Geometry == "snip2SameRect";
                var bottom = both ? Math.Min(rect.Width, rect.Height) * Math.Max(0f, Math.Min(0.5f, ShapeCurveGeometry.Adjustment(shape, 1, 0f))) : 0f;
                var path = new SKPath();
                path.MoveTo(rect.Left + (both ? top : 0f), rect.Top);
                path.LineTo(rect.Right - top, rect.Top);
                path.LineTo(rect.Right, rect.Top + top);
                path.LineTo(rect.Right, rect.Bottom - bottom);
                path.LineTo(rect.Right - bottom, rect.Bottom);
                path.LineTo(rect.Left + bottom, rect.Bottom);
                path.LineTo(rect.Left, rect.Bottom - bottom);
                path.LineTo(rect.Left, rect.Top + (both ? top : 0f));
                path.Close();
                return path;
            }

            return ShapeExtendedGeometry.Build(shape, rect);
        }

        private static SKPath BuildCallout(double[] adjusts, SKRect r)
        {
            var dx = adjusts != null && adjusts.Length > 0 ? (float)adjusts[0] : -0.20833f;
            var dy = adjusts != null && adjusts.Length > 1 ? (float)adjusts[1] : 0.625f;
            var tip = new SKPoint(r.MidX + r.Width * dx, r.MidY + r.Height * dy);
            var x1 = r.Left + r.Width * (dx > 0f ? 7f : 2f) / 12f;
            var x2 = r.Left + r.Width * (dx > 0f ? 10f : 5f) / 12f;
            var y1 = r.Top + r.Height * (dy > 0f ? 7f : 2f) / 12f;
            var y2 = r.Top + r.Height * (dy > 0f ? 10f : 5f) / 12f;
            var vertical = Math.Abs(dy) > Math.Abs(dx);
            var path = new SKPath();
            path.MoveTo(r.Left, r.Top);
            if (vertical && dy <= 0f)
            {
                path.LineTo(x1, r.Top); path.LineTo(tip); path.LineTo(x2, r.Top);
            }
            path.LineTo(r.Right, r.Top);
            if (!vertical && dx > 0f)
            {
                path.LineTo(r.Right, y1); path.LineTo(tip); path.LineTo(r.Right, y2);
            }
            path.LineTo(r.Right, r.Bottom);
            if (vertical && dy > 0f)
            {
                path.LineTo(x2, r.Bottom); path.LineTo(tip); path.LineTo(x1, r.Bottom);
            }
            path.LineTo(r.Left, r.Bottom);
            if (!vertical && dx <= 0f)
            {
                path.LineTo(r.Left, y2); path.LineTo(tip); path.LineTo(r.Left, y1);
            }
            path.Close();
            return path;
        }
    }
}
