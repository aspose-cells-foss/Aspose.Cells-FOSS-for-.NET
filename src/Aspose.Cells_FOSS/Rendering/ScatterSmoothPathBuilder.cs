using System.Collections.Generic;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ScatterSmoothPathBuilder
    {
        public static void Append(SKPath path, List<SKPoint> points)
        {
            for (var index = 0; index < points.Count - 1; index++)
            {
                var start = points[index];
                var end = points[index + 1];
                var before = index > 0 ? points[index - 1] : start;
                var after = index + 2 < points.Count ? points[index + 2] : end;
                var firstDivisor = index == 0 ? 3f : 6f;
                var secondDivisor = index + 2 == points.Count ? 3f : 6f;
                var control1 = new SKPoint(
                    start.X + (end.X - before.X) / firstDivisor,
                    start.Y + (end.Y - before.Y) / firstDivisor);
                var control2 = new SKPoint(
                    end.X - (after.X - start.X) / secondDivisor,
                    end.Y - (after.Y - start.Y) / secondDivisor);
                path.CubicTo(control1, control2, end);
            }
        }
    }
}
