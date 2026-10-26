using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class Column3DWallRenderer
    {
        public static void Draw(SKCanvas canvas, SKRect area, ParsedChart chart, SKPoint wallDepth)
        {
            if (canvas == null || chart == null
                || (chart.SideWallImageData == null && chart.BackWallImageData == null))
            {
                return;
            }

            var backBottomLeft = Column3DGeometry.GridPoint(area, 0d, 0d);
            var backBottomRight = Column3DGeometry.GridPoint(area, 1d, 0d);
            var backTopLeft = Column3DGeometry.GridPoint(area, 0d, 1d);
            var backTopRight = Column3DGeometry.GridPoint(area, 1d, 1d);
            var frontBottomLeft = Offset(backBottomLeft, wallDepth);
            var frontTopLeft = Offset(backTopLeft, wallDepth);
            Area3DWallFillRenderer.Draw(canvas, frontBottomLeft, frontTopLeft,
                backBottomLeft, backTopLeft, backBottomRight, backTopRight,
                chart.SideWallImageData, chart.BackWallImageData, SKColors.Transparent);
        }

        public static void DrawFloorGrid(SKCanvas canvas, SKRect area, ParsedChart chart,
            SKPoint wallDepth, SKPaint grid, int count)
        {
            var floorLeft = Column3DGeometry.GridPoint(area, 0d, 0d);
            var floorRight = Column3DGeometry.GridPoint(area, 1d, 0d);
            var frontLeft = Offset(floorLeft, wallDepth);
            var frontRight = Column3DGeometry.UsesSeriesDepth(chart)
                ? Column3DGeometry.ResolveSeriesFloorPoint(area, 1d, 1d)
                : Offset(floorRight, wallDepth);
            canvas.DrawLine(floorLeft, floorRight, grid);
            canvas.DrawLine(floorLeft, frontLeft, grid);
            canvas.DrawLine(frontLeft, frontRight, grid);

            for (var index = 0; index <= count; index++)
            {
                var ratio = count > 0 ? index / (double)count : 0d;
                var front = Column3DGeometry.UsesSeriesDepth(chart)
                    ? Column3DGeometry.ResolveSeriesFloorPoint(area, ratio, 1d)
                    : Offset(Interpolate(floorLeft, floorRight, ratio), wallDepth);
                canvas.DrawLine(front.X, front.Y, front.X, front.Y + area.Height * 0.012f, grid);
            }
        }

        private static SKPoint Offset(SKPoint point, SKPoint offset)
        {
            return new SKPoint(point.X + offset.X, point.Y + offset.Y);
        }

        private static SKPoint Interpolate(SKPoint start, SKPoint end, double ratio)
        {
            return new SKPoint(start.X + (end.X - start.X) * (float)ratio,
                start.Y + (end.Y - start.Y) * (float)ratio);
        }
    }
}
