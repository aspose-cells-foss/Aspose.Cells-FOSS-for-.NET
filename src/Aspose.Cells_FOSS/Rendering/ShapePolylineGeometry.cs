using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapePolylineGeometry
    {
        public static SKPoint[] Build(SKPoint start, SKPoint end, ParsedShape shape)
        {
            var a1 = ShapeCurveGeometry.Adjustment(shape, 0, 0.5f);
            var a2 = ShapeCurveGeometry.Adjustment(shape, 1, 0.5f);
            SKPoint[] points;
            switch (shape.Geometry)
            {
                case "bentConnector2":
                    points = new[] { new SKPoint(0f, 0f), new SKPoint(1f, 0f), new SKPoint(1f, 1f) };
                    break;
                case "bentConnector3":
                    points = new[] { new SKPoint(0f, 0f), new SKPoint(a1, 0f), new SKPoint(a1, 1f), new SKPoint(1f, 1f) };
                    break;
                case "bentConnector4":
                    points = new[] { new SKPoint(0f, 0f), new SKPoint(a1, 0f), new SKPoint(a1, a2), new SKPoint(1f, a2), new SKPoint(1f, 1f) };
                    break;
                default:
                    return new[] { start, end };
            }
            return ShapeCurveGeometry.Map(points, start, end, shape);
        }
    }
}
