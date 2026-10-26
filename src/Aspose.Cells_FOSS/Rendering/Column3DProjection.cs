using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    // Coordinates are category [0,1], value [0,1], and depth [0,1].
    // Depth zero is the front plane. All scene vertices use this camera.
    internal sealed class Column3DProjection
    {
        private readonly double _sinX;
        private readonly double _cosX;
        private readonly double _sinY;
        private readonly double _cosY;
        private readonly double _height;
        private readonly double _depth;
        private readonly double _perspective;
        private readonly SKRect _bounds;
        private readonly SKRect _cameraBounds;
        private readonly double _unitWidth;
        private readonly double _cameraOffsetX;
        private readonly double _cameraOffsetY;

        public Column3DProjection(SKRect chartArea, ParsedChart chart, int categoryCount)
            : this(chartArea, chart, categoryCount, 0f)
        {
        }

        public Column3DProjection(SKRect chartArea, ParsedChart chart, int categoryCount,
            float labelReserve)
        {
            var rx = chart.RotationXDeg * Math.PI / 180d;
            var ry = chart.RotationYDeg * Math.PI / 180d;
            _sinX = Math.Sin(rx);
            _cosX = Math.Cos(rx);
            _sinY = Math.Sin(ry);
            _cosY = Math.Cos(ry);
            _height = Math.Max(0.05d, (chartArea.Height - labelReserve) / chartArea.Width * 0.6d);
            _depth = Math.Max(1, chart.Series.Count) / (double)Math.Max(1, categoryCount)
                * Math.Max(0.01d, chart.DepthPercent / 100d);
            var perspective = chart.RightAngleAxes ? 0d : Math.Max(0d, chart.PerspectiveDeg);
            // The high-perspective Excel frame is an off-axis view: the camera's
            // vanishing point stays inside the plot instead of the near box corner.
            // Interpolate the framing correction from the default 30-degree view.
            var framing = Math.Max(0d, Math.Min(1d, (perspective - 30d) / 70d));
            _perspective = Math.Tan(Math.Min(120d, perspective) * Math.PI / 360d)
                / (0.48d + 0.29d * framing);
            _cameraOffsetX = -0.14d * framing;
            _cameraOffsetY = 0.0925d * framing;
            _unitWidth = chartArea.Width * 0.62d;
            // Keep the entire box in front of the camera even for extreme settings.
            var radius = Math.Abs(_cosX * _sinY) * 0.5d
                + Math.Abs(_cosX * _cosY) * _depth * 0.5d
                + Math.Abs(_sinX) * _height * 0.5d;
            _perspective = Math.Min(_perspective, 0.9d / Math.Max(0.001d, radius));
            _bounds = new SKRect(chartArea.Left + chartArea.Width * (0.158f + (float)framing * 0.006f),
                chartArea.Top + chartArea.Height * (0.19f + (float)framing * 0.011f),
                chartArea.Left + chartArea.Width * (0.85f - (float)framing * 0.054f),
                chartArea.Top + chartArea.Height * 0.752f - labelReserve * 0.34f);

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
                    / _cameraBounds.Width * _bounds.Width,
                _bounds.Top + (point.Y - _cameraBounds.Top)
                    / _cameraBounds.Height * _bounds.Height);
        }

        public double CameraDepth(double category, double value, double depth)
        {
            var yawDepth = -_sinY * (category - 0.5d)
                + _cosY * (depth - 0.5d) * _depth;
            return _cosX * yawDepth - _sinX * (value - 0.5d) * _height;
        }

        public static double ColumnWidth(ParsedChart chart, int categoryCount)
        {
            return 1d / (Math.Max(1, categoryCount)
                * (1d + Math.Max(0d, chart.GapWidthPercent) / 100d));
        }

        public static double ColumnDepth(ParsedChart chart)
        {
            return 1d / (Math.Max(1, chart.Series.Count)
                * (1d + Math.Max(0d, chart.GapDepthPercent) / 100d));
        }

        public SKSize WallSize(bool side)
        {
            return new SKSize((float)(_unitWidth * (side ? _depth : 1d)),
                (float)(_unitWidth * _height));
        }

        public SKMatrix WallTransform(bool side)
        {
            var a = side ? Project(0, 1, 0) : Project(0, 1, 1);
            var b = side ? Project(0, 1, 1) : Project(1, 1, 1);
            var c = side ? Project(0, 0, 1) : Project(1, 0, 1);
            var d = side ? Project(0, 0, 0) : Project(0, 0, 1);
            var dx1 = b.X - c.X;
            var dx2 = d.X - c.X;
            var dy1 = b.Y - c.Y;
            var dy2 = d.Y - c.Y;
            var dx3 = a.X - b.X + c.X - d.X;
            var dy3 = a.Y - b.Y + c.Y - d.Y;
            var determinant = dx1 * dy2 - dx2 * dy1;
            if (Math.Abs(determinant) < 0.00001f)
            {
                return SKMatrix.CreateIdentity();
            }
            var g = (dx3 * dy2 - dx2 * dy3) / determinant;
            var h = (dx1 * dy3 - dx3 * dy1) / determinant;
            var size = WallSize(side);
            return new SKMatrix
            {
                ScaleX = (b.X - a.X + g * b.X) / size.Width,
                SkewX = (d.X - a.X + h * d.X) / size.Height,
                TransX = a.X,
                SkewY = (b.Y - a.Y + g * b.Y) / size.Width,
                ScaleY = (d.Y - a.Y + h * d.Y) / size.Height,
                TransY = a.Y,
                Persp0 = g / size.Width, Persp1 = h / size.Height, Persp2 = 1f
            };
        }

        public SKPoint[] ColumnCorners(double center, double seriesCenter,
            double width, double depth, double bottom, double top)
        {
            var x0 = center - width * 0.5d;
            var x1 = center + width * 0.5d;
            var z0 = seriesCenter - depth * 0.5d;
            var z1 = seriesCenter + depth * 0.5d;
            return new SKPoint[]
            {
                Project(x0, bottom, z0), Project(x1, bottom, z0),
                Project(x1, bottom, z1), Project(x0, bottom, z1),
                Project(x0, top, z0), Project(x1, top, z0),
                Project(x1, top, z1), Project(x0, top, z1)
            };
        }

        private SKPoint CameraPoint(double category, double value, double depth)
        {
            var x = category - 0.5d;
            var y = (value - 0.5d) * _height;
            var z = (depth - 0.5d) * _depth;
            var yawDepth = -_sinY * x + _cosY * z;
            var denominator = 1d + _perspective * CameraDepth(category, value, depth);
            return new SKPoint((float)((_cosY * x + _sinY * z + _cameraOffsetX) / denominator),
                (float)(-(_cosX * y + _sinX * yawDepth + _cameraOffsetY) / denominator));
        }
    }
}
