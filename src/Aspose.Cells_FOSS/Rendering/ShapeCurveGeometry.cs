using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapeCurveGeometry
    {
        public static SKPoint[] Build(SKPoint start, SKPoint end, ParsedShape shape)
        {
            if (shape.Geometry == "curvedConnector4")
            {
                var x2 = Adjustment(shape, 0, 0.5f);
                var y4 = Adjustment(shape, 1, 0.5f);
                var x3 = (1f + x2) / 2f;
                return Map(new[] { new SKPoint(0f, 0f), new SKPoint(x2 / 2f, 0f),
                    new SKPoint(x2, y4 / 4f), new SKPoint(x2, y4 / 2f),
                    new SKPoint(x2, y4 * 0.75f), new SKPoint((x2 + x3) / 2f, y4), new SKPoint(x3, y4),
                    new SKPoint((x3 + 1f) / 2f, y4), new SKPoint(1f, (1f + y4) / 2f), new SKPoint(1f, 1f) }, start, end, shape);
            }
            if (shape.Geometry != "curvedConnector3")
            {
                return new[] { start, new SKPoint((start.X + end.X) / 2f, start.Y),
                    new SKPoint(end.X, (start.Y + end.Y) / 2f), end };
            }

            // The three-leg preset consists of two cubics, with horizontal tangents
            // at both ends and a vertical tangent at the adjustable middle joint.
            var adj = shape.Adjusts != null && shape.Adjusts.Length > 0 ? (float)shape.Adjusts[0] : 0.5f;
            var normalized = new[] { new SKPoint(0f, 0f), new SKPoint(adj / 2f, 0f),
                new SKPoint(adj, 0.25f), new SKPoint(adj, 0.5f), new SKPoint(adj, 0.75f),
                new SKPoint((1f + adj) / 2f, 1f), new SKPoint(1f, 1f) };
            return Map(normalized, start, end, shape);
        }

        public static float Adjustment(ParsedShape shape, int index, float fallback)
        {
            return shape.Adjusts != null && shape.Adjusts.Length > index ? (float)shape.Adjusts[index] : fallback;
        }

        public static SKPoint[] Map(SKPoint[] normalized, SKPoint start, SKPoint end, ParsedShape shape)
        {
            var vertical = Math.Abs(Math.Sin(shape.RotationDeg * Math.PI / 180d)) > 0.707d;
            for (var i = 0; i < normalized.Length; i++)
            {
                var x = vertical ? normalized[i].Y : normalized[i].X;
                var y = vertical ? normalized[i].X : normalized[i].Y;
                normalized[i] = new SKPoint(start.X + (end.X - start.X) * x, start.Y + (end.Y - start.Y) * y);
            }

            return normalized;
        }
    }
}
