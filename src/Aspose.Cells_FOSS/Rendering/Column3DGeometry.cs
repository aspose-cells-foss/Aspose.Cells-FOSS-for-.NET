using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class Column3DGeometry
    {
        private const float CylinderFloorExpansion = 1.35f;
        private const float PercentStackedBoxFloorExpansion = 1.3f;
        private const float SeriesWallLeftDepthXRatio = 0.196f;
        private const float SeriesWallLeftDepthYRatio = 0.089f;
        private const float SeriesWallRightDepthXRatio = 0.0165f;
        private const float SeriesWallRightDepthYRatio = 0.158f;

        public static bool IsPercentStackedBox(ParsedChart chart)
        {
            return chart != null && chart.IsPercentStacked
                && string.Equals(chart.BarShape, "box", StringComparison.Ordinal);
        }

        public static bool UsesSeriesDepth(ParsedChart chart)
        {
            return chart != null && chart.IsBar3D && !chart.IsStacked
                && !chart.IsPercentStacked && chart.HasVisibleSeriesAxis;
        }

        public static SKPoint ResolveFloorDepth(SKRect area, ParsedChart chart,
            SKPoint defaultDepth, float barWidth, SKPoint shapeDepth, float horizontalSlope)
        {
            if (chart == null)
            {
                return defaultDepth;
            }

            if (UsesSeriesDepth(chart))
            {
                return ResolveSeriesWallDepth(area, 0d);
            }

            if (string.Equals(chart.BarShape, "cylinder", StringComparison.Ordinal))
            {
                var widthAxisY = barWidth * horizontalSlope * 0.5f;
                var depthAxisY = shapeDepth.Y * 0.5f;
                var verticalRadius = (float)Math.Sqrt(
                    widthAxisY * widthAxisY + depthAxisY * depthAxisY);
                var floorHeight = (-shapeDepth.Y + verticalRadius * 2f + 0.5f)
                    * CylinderFloorExpansion;
                return new SKPoint(Math.Min(defaultDepth.X, -floorHeight * 2f),
                    Math.Max(defaultDepth.Y, floorHeight));
            }

            if (!IsPercentStackedBox(chart))
            {
                return defaultDepth;
            }

            var projectedWidthHeight = Math.Abs(barWidth * horizontalSlope);
            var boxFloorHeight = (-shapeDepth.Y + projectedWidthHeight + 0.5f)
                * PercentStackedBoxFloorExpansion;
            return new SKPoint(Math.Min(defaultDepth.X, -boxFloorHeight * 2f),
                Math.Max(defaultDepth.Y, boxFloorHeight));
        }

        public static SKPoint ResolveBoxBottomCenterOffset(SKPoint floorMidpoint, SKPoint depth)
        {
            return new SKPoint(floorMidpoint.X - depth.X * 0.5f,
                floorMidpoint.Y - depth.Y * 0.5f);
        }

        public static SKPoint ResolveSeriesFloorOffset(SKPoint floorDepth,
            int seriesIndex, int seriesCount)
        {
            var ratio = seriesCount > 1
                ? 0.85f - 0.60f * seriesIndex / (seriesCount - 1)
                : 0.5f;
            return new SKPoint(floorDepth.X * ratio, floorDepth.Y * ratio);
        }

        public static SKPoint ResolveSeriesWallDepth(SKRect area, double horizontalRatio)
        {
            var left = new SKPoint(-area.Width * SeriesWallLeftDepthXRatio,
                area.Height * SeriesWallLeftDepthYRatio);
            var right = new SKPoint(-area.Width * SeriesWallRightDepthXRatio,
                area.Height * SeriesWallRightDepthYRatio);
            return Interpolate(left, right, horizontalRatio);
        }

        public static SKPoint ResolveSeriesFloorPoint(SKRect area, double horizontalRatio,
            double depthRatio)
        {
            var back = GridPoint(area, horizontalRatio, 0d);
            var depth = ResolveSeriesWallDepth(area, horizontalRatio);
            return new SKPoint(back.X + depth.X * (float)depthRatio,
                back.Y + depth.Y * (float)depthRatio);
        }

        public static SKPoint ResolveSeriesShapeDepth(SKRect area, float barWidth,
            double horizontalRatio)
        {
            var wallDepth = ResolveSeriesWallDepth(area, horizontalRatio);
            var wallLength = (float)Math.Sqrt(
                wallDepth.X * wallDepth.X + wallDepth.Y * wallDepth.Y);
            if (wallLength < 0.0001f)
            {
                return new SKPoint(barWidth * 0.85f, 0f);
            }
            var scale = barWidth * 0.85f / wallLength;
            return new SKPoint(-wallDepth.X * scale, -wallDepth.Y * scale);
        }

        public static SKPoint GridPoint(SKRect area, double horizontalRatio, double valueRatio)
        {
            var bottomLeft = new SKPoint(area.Left + area.Width * 0.248f,
                area.Top + area.Height * 0.593f);
            var bottomRight = new SKPoint(area.Left + area.Width * 0.838f,
                area.Top + area.Height * 0.738f);
            var topLeft = new SKPoint(area.Left + area.Width * 0.233f,
                area.Top + area.Height * 0.193f);
            var topRight = new SKPoint(area.Left + area.Width * 0.860f,
                area.Top + area.Height * 0.246f);
            var bottom = Interpolate(bottomLeft, bottomRight, horizontalRatio);
            var top = Interpolate(topLeft, topRight, horizontalRatio);
            return Interpolate(bottom, top, valueRatio);
        }

        private static SKPoint Interpolate(SKPoint start, SKPoint end, double ratio)
        {
            return new SKPoint(
                start.X + (end.X - start.X) * (float)ratio,
                start.Y + (end.Y - start.Y) * (float)ratio);
        }

        public static float ResolveBoxHeightExtension(SKPoint bottomCenterOffset, SKPoint depth)
        {
            return Math.Max(0f, bottomCenterOffset.Y + depth.Y - 2.5f);
        }

        public static void ResolveShapeOffsets(ParsedChart chart, float barWidth,
            SKPoint depth, float horizontalSlope, ref SKPoint floorOffset,
            out float heightExtension)
        {
            heightExtension = 0f;
            if (string.Equals(chart.BarShape, "cylinder", StringComparison.Ordinal))
            {
                floorOffset = Column3DCylinderRenderer.ResolveBottomCenterOffset(
                    floorOffset, depth);
                heightExtension = Column3DCylinderRenderer.ResolveHeightExtension(
                    floorOffset.Y, barWidth, depth, horizontalSlope);
            }
            else if (IsPercentStackedBox(chart))
            {
                floorOffset = ResolveBoxBottomCenterOffset(floorOffset, depth);
                heightExtension = ResolveBoxHeightExtension(floorOffset, depth);
            }
        }
    }
}
