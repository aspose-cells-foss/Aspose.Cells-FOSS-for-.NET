using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapePresetDetails
    {
        public static bool DrawFill(SKCanvas canvas, SKRect r, ParsedShape shape)
        {
            if (shape.Geometry != "bevel") return false;
            var inset = BevelInset(shape, r);
            var x1 = r.Left + inset;
            var x2 = r.Right - inset;
            var y1 = r.Top + inset;
            var y2 = r.Bottom - inset;
            // Disjoint faces: do not paint a base rectangle under translucent facets.
            Face(canvas, shape.FillColor, new float[] { x1,y1, x2,y1, x2,y2, x1,y2 });
            Face(canvas, FaceColor(shape.FillColor, 0.2f), new float[] { r.Left,r.Top, r.Right,r.Top, x2,y1, x1,y1 });
            Face(canvas, FaceColor(shape.FillColor, -0.2f), new float[] { r.Left,r.Bottom, x1,y2, x2,y2, r.Right,r.Bottom });
            Face(canvas, FaceColor(shape.FillColor, 0.4f), new float[] { r.Left,r.Top, x1,y1, x1,y2, r.Left,r.Bottom });
            Face(canvas, FaceColor(shape.FillColor, -0.4f), new float[] { r.Right,r.Top, r.Right,r.Bottom, x2,y2, x2,y1 });
            return true;
        }

        public static void DrawLines(SKCanvas canvas, SKRect r, ParsedShape shape)
        {
            if (!shape.HasLine) return;
            using (var path = new SKPath())
            using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, Color = shape.LineColor,
                StrokeWidth = (float)Math.Max(0.5d, shape.LineWidthPt), IsAntialias = true, StrokeJoin = SKStrokeJoin.Miter })
            {
                if (shape.Geometry == "flowChartMagneticDrum")
                {
                    path.ArcTo(new SKRect(r.Right - r.Width / 3f, r.Top, r.Right, r.Bottom), 90f, 180f, true);
                }
                else if (shape.Geometry == "accentBorderCallout3")
                {
                    var defaults = new float[] { 0.1875f, -0.08333f, 0.1875f, -0.16667f, 1f, -0.16667f, 1.12963f, -0.08333f };
                    var points = new SKPoint[4];
                    for (var i = 0; i < points.Length; i++)
                    {
                        points[i] = new SKPoint(r.Left + r.Width * ShapeExtendedGeometry.Adjustment(shape, i * 2 + 1, defaults[i * 2 + 1]),
                            r.Top + r.Height * ShapeExtendedGeometry.Adjustment(shape, i * 2, defaults[i * 2]));
                    }
                    path.MoveTo(points[0].X, r.Top);
                    path.LineTo(points[0].X, r.Bottom);
                    path.MoveTo(points[0]);
                    for (var i = 1; i < points.Length; i++) path.LineTo(points[i]);
                }
                else if (shape.Geometry == "bevel")
                {
                    var inset = BevelInset(shape, r);
                    var inner = new SKRect(r.Left + inset, r.Top + inset, r.Right - inset, r.Bottom - inset);
                    path.AddRect(inner);
                    path.MoveTo(r.Left, r.Top); path.LineTo(inner.Left, inner.Top);
                    path.MoveTo(r.Right, r.Top); path.LineTo(inner.Right, inner.Top);
                    path.MoveTo(r.Right, r.Bottom); path.LineTo(inner.Right, inner.Bottom);
                    path.MoveTo(r.Left, r.Bottom); path.LineTo(inner.Left, inner.Bottom);
                }
                else return;
                canvas.DrawPath(path, paint);
            }
        }

        private static float BevelInset(ParsedShape shape, SKRect r)
        {
            return Math.Min(r.Width, r.Height) * Math.Max(0f, Math.Min(0.5f, ShapeExtendedGeometry.Adjustment(shape, 0, 0.125f)));
        }

        private static void Face(SKCanvas canvas, SKColor color, float[] points)
        {
            using (var path = ShapeExtendedGeometry.Polygon(points))
            using (var paint = new SKPaint { Color = color, IsAntialias = true }) canvas.DrawPath(path, paint);
        }

        private static SKColor FaceColor(SKColor color, float amount)
        {
            return new SKColor(Channel(color.Red, amount), Channel(color.Green, amount), Channel(color.Blue, amount), color.Alpha);
        }

        private static byte Channel(byte value, float amount)
        {
            return (byte)Math.Round(amount >= 0f ? value + (255 - value) * amount : value * (1f + amount));
        }
    }
}
