using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    // Category, value and depth coordinates all use the same perspective camera.
    // Ribbon width is a depth interval, never a screen-space stroke width.
    internal sealed class Line3DProjection
    {
        private readonly double _sinX;
        private readonly double _cosX;
        private readonly double _sinY;
        private readonly double _cosY;
        private readonly double _height;
        private readonly double _depth;
        private readonly double _perspective;
        private readonly double _cameraOffsetX;
        private readonly double _cameraOffsetY;
        private readonly int _seriesCount;
        private readonly double _ribbonDepth;
        private readonly SKRect _bounds;
        private readonly SKRect _cameraBounds;

        public Line3DProjection(SKRect plot, ParsedChart chart, int categoryCount)
        {
            var rx = chart.RotationXDeg * Math.PI / 180d;
            var ry = chart.RotationYDeg * Math.PI / 180d;
            _sinX = Math.Sin(rx);
            _cosX = Math.Cos(rx);
            _sinY = Math.Sin(ry);
            _cosY = Math.Cos(ry);
            _seriesCount = Math.Max(1, chart.Series.Count);
            _depth = _seriesCount / (double)Math.Max(1, categoryCount)
                * Math.Max(0.01d, chart.DepthPercent / 100d);
            // A deep series box must not flatten the value dimension or move
            // its near plane towards the camera singularity. Normalize both
            // height and viewing distance by the horizontal extent of the box.
            var extent = Math.Max(1d, _depth);
            _height = Math.Max(0.05d, plot.Height / Math.Max(1f, plot.Width) * 0.76d)
                * Math.Max(1d, _depth * 0.875d);
            _ribbonDepth = 1d / (_seriesCount * (1d + Math.Max(0d, chart.GapDepthPercent) / 100d));
            var framing = chart.RightAngleAxes ? 0d
                : Math.Max(0d, Math.Min(1d, (chart.PerspectiveDeg - 30d) / 30d));
            _perspective = chart.RightAngleAxes ? 0d
                : Math.Tan(Math.Min(120d, Math.Max(0d, chart.PerspectiveDeg)) * Math.PI / 360d)
                    / (0.475d * extent * (1d + 0.22d * framing));
            // Excel's stronger perspective is an off-axis view, with the
            // vanishing point inside the plot rather than above its near edge.
            _cameraOffsetX = 0.21d * framing;
            _cameraOffsetY = -0.26d * framing;
            var radius = Math.Abs(_cosX * _sinY) * 0.5d
                + Math.Abs(_cosX * _cosY) * _depth * 0.5d
                + Math.Abs(_sinX) * _height * 0.5d;
            _perspective = Math.Min(_perspective, 0.9d / Math.Max(0.001d, radius));

            // Reserve the label bands around the complete projected box, not
            // around independently positioned walls and data series.
            _bounds = new SKRect(plot.Left + plot.Width * (0.048f + (float)framing * 0.109f),
                plot.Top + plot.Height * (0.045f - (float)framing * 0.072f),
                plot.Left + plot.Width * (0.876f - (float)framing * 0.131f),
                plot.Top + plot.Height * (0.965f - (float)framing * 0.05f));
            var minX = float.MaxValue;
            var minY = float.MaxValue;
            var maxX = float.MinValue;
            var maxY = float.MinValue;
            for (var x = 0; x <= 1; x++)
            {
                for (var y = 0; y <= 1; y++)
                {
                    for (var z = 0; z <= 1; z++)
                    {
                        var point = CameraPoint(x, y, z);
                        minX = Math.Min(minX, point.X);
                        minY = Math.Min(minY, point.Y);
                        maxX = Math.Max(maxX, point.X);
                        maxY = Math.Max(maxY, point.Y);
                    }
                }
            }
            _cameraBounds = new SKRect(minX, minY, maxX, maxY);
        }

        public SKPoint Project(double category, double value, double depth)
        {
            var point = CameraPoint(category, value, depth);
            return new SKPoint(_bounds.Left + (point.X - _cameraBounds.Left)
                    / Math.Max(0.000001f, _cameraBounds.Width) * _bounds.Width,
                _bounds.Top + (point.Y - _cameraBounds.Top)
                    / Math.Max(0.000001f, _cameraBounds.Height) * _bounds.Height);
        }

        public double SeriesDepth(int index)
        {
            return (index + 0.5d) / _seriesCount;
        }

        public SKPoint[] RibbonCorners(SKPoint start, SKPoint end, int seriesIndex)
        {
            var center = SeriesDepth(seriesIndex);
            var front = center - _ribbonDepth * 0.5d;
            var back = center + _ribbonDepth * 0.5d;
            return new SKPoint[]
            {
                Project(start.X, start.Y, front), Project(end.X, end.Y, front),
                Project(end.X, end.Y, back), Project(start.X, start.Y, back)
            };
        }

        public double RibbonShade(SKPoint start, SKPoint end, SKPoint[] corners)
        {
            var dx = end.X - start.X;
            var dy = (end.Y - start.Y) * _height;
            var length = Math.Sqrt(dx * dx + dy * dy);
            if (length < 0.000001d)
            {
                return 0.76d;
            }

            var edgeX = corners[1].X - corners[0].X;
            var edgeY = corners[1].Y - corners[0].Y;
            var depthX = corners[3].X - corners[0].X;
            var depthY = corners[3].Y - corners[0].Y;
            // A steep rising segment can expose the underside of the ribbon.
            // The projected winding distinguishes it from an upward-facing face.
            if (edgeX * depthY - edgeY * depthX > 0d)
            {
                return 0.40d;
            }

            var normalX = -dy / length;
            return Math.Max(0.4d, Math.Min(0.81d, 0.76d + 0.25d * normalX));
        }

        private SKPoint CameraPoint(double category, double value, double depth)
        {
            var x = category - 0.5d;
            var y = (value - 0.5d) * _height;
            var z = (depth - 0.5d) * _depth;
            var yawDepth = -_sinY * x + _cosY * z;
            var cameraDepth = _cosX * yawDepth - _sinX * y;
            var divisor = 1d + _perspective * cameraDepth;
            return new SKPoint((float)((_cosY * x + _sinY * z + _cameraOffsetX) / divisor),
                (float)((-(_cosX * y + _sinX * yawDepth) + _cameraOffsetY) / divisor));
        }
    }
}
