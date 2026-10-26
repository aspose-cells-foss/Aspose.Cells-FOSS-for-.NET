using System;
using System.Globalization;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class Bar3DChartRenderer
    {
        private static readonly SKColor GridColor = new SKColor(0xD9, 0xD9, 0xD9);
        private const float LabelSizePt = 9f;
        private readonly RenderContext _context;
        private readonly DateSystem _dateSystem;
        private readonly CultureInfo _culture;

        public Bar3DChartRenderer(RenderContext context, DateSystem dateSystem, CultureInfo culture)
        {
            _context = context;
            _dateSystem = dateSystem;
            _culture = culture ?? CultureInfo.CurrentCulture;
        }

        public void Draw(SKCanvas canvas, SKRect area, SKRect plot, ParsedChart chart, double min, double max, double major)
        {
            if (canvas == null || chart == null || chart.Series.Count == 0 || max <= min)
            {
                return;
            }

            var denseGrid = chart.HasMinorValueGridlines;
            var topLeftX = 0.086f;
            var topLeftY = -0.018f;
            var topRightX = 0.168f;
            var topRightY = 0.076f;
            var bottomLeftX = 0.104f;
            var bottomLeftY = 0.512f;
            var bottomRightX = 0.188f;
            var bottomRightY = 0.743f;
            if (UsesPerspectiveFan(chart))
            {
                topLeftX = 0.174f;
                topLeftY = 0.141f;
                topRightX = 0.169f;
                topRightY = -0.007f;
                bottomLeftX = 0.176f;
                bottomLeftY = 0.746f;
                bottomRightX = 0.177f;
                bottomRightY = 0.912f;
            }
            else if (chart.HasDateCategoryAxis && !chart.RightAngleAxes
                && chart.PerspectiveDeg >= 100d)
            {
                topLeftX = 0.079f;
                topLeftY = 0.166f;
                topRightX = 0.231f;
                topRightY = 0.071f;
                bottomLeftX = 0.116f;
                bottomLeftY = 0.725f;
                bottomRightX = 0.284f;
                bottomRightY = 0.925f;
            }
            else if (chart.HasDateCategoryAxis)
            {
                topLeftX = 0.079f;
                topLeftY = 0.071f;
                topRightX = 0.133f;
                topRightY = 0.168f;
                bottomLeftX = 0.097f;
                bottomLeftY = 0.760f;
                bottomRightX = 0.153f;
                bottomRightY = 1.101f;
            }
            else if (denseGrid)
            {
                topLeftX = 0.073f;
                topLeftY = -0.077f;
                topRightX = 0.116f;
                topRightY = 0.023f;
                bottomLeftX = 0.094f;
                bottomLeftY = 0.746f;
                bottomRightX = 0.143f;
                bottomRightY = 1.002f;
            }
            else if (chart.IsPercentStacked)
            {
                topLeftX = 0.145f;
                topLeftY = -0.011f;
                topRightX = 0.117f;
                topRightY = 0.082f;
                bottomLeftX = 0.161f;
                bottomLeftY = 0.674f;
                bottomRightX = 0.143f;
                bottomRightY = 0.916f;
            }
            else if (chart.IsStacked)
            {
                topLeftX = 0.080f;
                topLeftY = -0.011f;
                topRightX = 0.143f;
                topRightY = 0.082f;
                bottomLeftX = 0.100f;
                bottomLeftY = 0.674f;
                bottomRightX = 0.165f;
                bottomRightY = 0.916f;
            }

            var topLeft = new SKPoint(
                plot.Left + plot.Width * topLeftX,
                plot.Top + plot.Height * topLeftY);
            var topRight = new SKPoint(
                plot.Right - plot.Width * topRightX,
                plot.Top + plot.Height * topRightY);
            var bottomLeft = new SKPoint(
                plot.Left + plot.Width * bottomLeftX,
                plot.Top + plot.Height * bottomLeftY);
            var bottomRight = new SKPoint(
                plot.Right - plot.Width * bottomRightX,
                plot.Top + plot.Height * bottomRightY);
            SKPoint wallFrontTopLeft;
            SKPoint wallFrontBottomLeft;
            ResolveSideWall(plot, topLeft, bottomLeft, out wallFrontTopLeft, out wallFrontBottomLeft);
            var sideWallImage = chart.HasDateCategoryAxis ? null : chart.SideWallImageData;
            var wallFallback = chart.HasDateCategoryAxis ? SKColors.Transparent : SKColors.White;
            Area3DWallFillRenderer.Draw(
                canvas,
                wallFrontBottomLeft,
                wallFrontTopLeft,
                bottomLeft,
                topLeft,
                bottomRight,
                topRight,
                sideWallImage,
                chart.BackWallImageData,
                wallFallback);
            if (chart.HasDateCategoryAxis)
            {
                var floorOffset = new SKPoint(
                    wallFrontBottomLeft.X - bottomLeft.X,
                    wallFrontBottomLeft.Y - bottomLeft.Y);
                var floorFrontRight = Offset(bottomRight, floorOffset);
                Area3DWallFillRenderer.DrawFloor(
                    canvas,
                    bottomLeft,
                    bottomRight,
                    floorFrontRight,
                    wallFrontBottomLeft,
                    chart.BackWallImageData,
                    new SKColor(0xE5, 0xD4, 0xA2));
            }
            DrawSideWall(canvas, wallFrontTopLeft, topLeft, bottomLeft, wallFrontBottomLeft);
            var renderedMajor = denseGrid || chart.HasDateCategoryAxis
                ? major
                : NiceNumber((max - min) / 4d);
            DrawGridAndValueAxis(canvas, area, topLeft, topRight, bottomLeft, bottomRight,
                wallFrontTopLeft, wallFrontBottomLeft, chart, min, max, renderedMajor);
            DrawBars(canvas, topLeft, topRight, bottomLeft, bottomRight,
                wallFrontTopLeft, wallFrontBottomLeft, chart, min, max);
            DrawCategoryLabels(canvas, area, wallFrontTopLeft, wallFrontBottomLeft, chart);
        }

        private static bool UsesPerspectiveFan(ParsedChart chart)
        {
            return chart.IsPercentStacked
                && !chart.RightAngleAxes
                && Math.Abs(chart.RotationXDeg) < 0.01d
                && chart.PerspectiveDeg > 0d;
        }

        private void DrawGridAndValueAxis(SKCanvas canvas, SKRect area,
            SKPoint topLeft, SKPoint topRight, SKPoint bottomLeft, SKPoint bottomRight,
            SKPoint wallFrontTopLeft, SKPoint wallFrontBottomLeft,
            ParsedChart chart, double min, double max, double major)
        {
            DrawMinorGridlines(canvas, topLeft, topRight, bottomLeft, bottomRight, chart, min, max, major);
            using (var grid = new SKPaint { Style = SKPaintStyle.Stroke, Color = GridColor, StrokeWidth = 0.75f, IsAntialias = true })
            {
                var steps = (int)Math.Round((max - min) / major);
                for (var i = 0; i <= steps; i++)
                {
                    var value = min + i * major;
                    var ratio = (value - min) / (max - min);
                    var projectedRatio = ProjectValueRatio((float)ratio);
                    var top = Interpolate(topLeft, topRight, projectedRatio);
                    var bottom = Interpolate(bottomLeft, bottomRight, projectedRatio);
                    canvas.DrawLine(top, bottom, grid);
                    if (chart.HasDateCategoryAxis && chart.PerspectiveDeg >= 100d && i % 2 != 0)
                    {
                        continue;
                    }

                    var label = ChartXmlParser.FormatLabel(
                        value / chart.ValueDisplayUnitDivisor,
                        chart.ValueFormatCode,
                        _dateSystem,
                        _culture);
                    DrawCenteredText(canvas, area, label != null ? label.TrimEnd() : string.Empty,
                        bottom.X, bottom.Y + 11f, chart.AxisLabelColor);
                }

                var categoryTicks = chart.HasMinorValueGridlines ? 6 : 3;
                for (var i = 0; i <= categoryTicks; i++)
                {
                    var tick = Interpolate(wallFrontTopLeft, wallFrontBottomLeft, i / (float)categoryTicks);
                    canvas.DrawLine(tick.X, tick.Y, tick.X - 2.8f, tick.Y, grid);
                }
            }

            DrawValueAxisLine(canvas, bottomLeft, bottomRight, chart, min, max, major);

            if (!string.IsNullOrEmpty(chart.ValueDisplayUnitLabel))
            {
                DrawText(canvas, area, chart.ValueDisplayUnitLabel,
                    bottomRight.X - 5f, bottomRight.Y + 27f, chart.AxisLabelColor, true);
            }
        }

        private static void DrawSideWall(SKCanvas canvas,
            SKPoint frontTopLeft, SKPoint backTopLeft, SKPoint backBottomLeft, SKPoint frontBottomLeft)
        {
            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = GridColor,
                StrokeWidth = 0.75f,
                IsAntialias = true
            })
            {
                canvas.DrawLine(frontTopLeft, backTopLeft, paint);
                canvas.DrawLine(frontTopLeft, frontBottomLeft, paint);
                canvas.DrawLine(frontBottomLeft, backBottomLeft, paint);
            }
        }

        private static void ResolveSideWall(SKRect plot, SKPoint backTopLeft, SKPoint backBottomLeft,
            out SKPoint frontTopLeft, out SKPoint frontBottomLeft)
        {
            frontTopLeft = new SKPoint(
                backTopLeft.X - Math.Max(4f, plot.Width * 0.022f),
                backTopLeft.Y + plot.Height * 0.010f);
            frontBottomLeft = new SKPoint(
                backBottomLeft.X - Math.Max(3f, plot.Width * 0.014f),
                backBottomLeft.Y + plot.Height * 0.025f);
        }

        private static void DrawMinorGridlines(SKCanvas canvas,
            SKPoint topLeft, SKPoint topRight, SKPoint bottomLeft, SKPoint bottomRight,
            ParsedChart chart, double min, double max, double major)
        {
            if (!chart.HasMinorValueGridlines || major <= 0d)
            {
                return;
            }

            var minor = chart.MinorUnit ?? major / 5d;
            if (minor <= 0d)
            {
                return;
            }

            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.MinorValueGridlineColor,
                StrokeWidth = (float)chart.MinorValueGridlineWidthPt,
                IsAntialias = true
            })
            {
                var steps = (int)Math.Round((max - min) / minor);
                for (var i = 1; i < steps; i++)
                {
                    var value = min + i * minor;
                    var majorRatio = (value - min) / major;
                    if (Math.Abs(majorRatio - Math.Round(majorRatio)) < 0.0001d)
                    {
                        continue;
                    }

                    var ratio = ProjectValueRatio((float)((value - min) / (max - min)));
                    canvas.DrawLine(
                        Interpolate(topLeft, topRight, ratio),
                        Interpolate(bottomLeft, bottomRight, ratio),
                        paint);
                }
            }
        }

        private static void DrawValueAxisLine(SKCanvas canvas, SKPoint bottomLeft, SKPoint bottomRight,
            ParsedChart chart, double min, double max, double major)
        {
            if (!chart.HasValueAxisLine)
            {
                return;
            }

            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.ValueAxisLineColor,
                StrokeWidth = (float)chart.ValueAxisLineWidthPt,
                IsAntialias = true
            })
            {
                canvas.DrawLine(bottomLeft, bottomRight, paint);
                if (string.Equals(chart.ValueAxisMajorTickMark, "none", StringComparison.Ordinal))
                {
                    return;
                }

                var steps = (int)Math.Round((max - min) / major);
                for (var i = 0; i <= steps; i++)
                {
                    var ratio = ProjectValueRatio(i / (float)steps);
                    var tick = Interpolate(bottomLeft, bottomRight, ratio);
                    var startY = tick.Y;
                    var endY = tick.Y + 5.6f;
                    if (string.Equals(chart.ValueAxisMajorTickMark, "cross", StringComparison.Ordinal))
                    {
                        startY -= 2.8f;
                        endY -= 2.8f;
                    }
                    else if (string.Equals(chart.ValueAxisMajorTickMark, "in", StringComparison.Ordinal))
                    {
                        endY = tick.Y - 5.6f;
                    }
                    canvas.DrawLine(tick.X, startY, tick.X, endY, paint);
                }
            }
        }

        private void DrawBars(SKCanvas canvas,
            SKPoint topLeft, SKPoint topRight, SKPoint bottomLeft, SKPoint bottomRight,
            SKPoint wallFrontTopLeft, SKPoint wallFrontBottomLeft,
            ParsedChart chart, double min, double max)
        {
            var pointCount = chart.Categories.Count;
            if (pointCount <= 0)
            {
                return;
            }

            var leftHeight = Distance(topLeft, bottomLeft);
            var categorySlotCount = ResolveCategorySlotCount(chart, pointCount);
            var slot = leftHeight / categorySlotCount;
            var thicknessDirection = ResolveBarThicknessDirection(topLeft, bottomLeft);
            if (chart.IsStacked || chart.IsPercentStacked)
            {
                DrawStackedBars(canvas, topLeft, topRight, bottomLeft, bottomRight,
                    wallFrontTopLeft, wallFrontBottomLeft,
                    thicknessDirection, chart, min, max, slot);
                return;
            }

            var barWidth = ResolveClusteredBarWidth(chart, slot);
            for (var i = 0; i < pointCount; i++)
            {
                var categoryIndex = pointCount - 1 - i;
                for (var s = chart.Series.Count - 1; s >= 0; s--)
                {
                    var series = chart.Series[s];
                    if (i >= series.Values.Count || !series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var ratio = (series.Values[i].Value - min) / (max - min);
                    var seriesSpacing = ResolveClusteredSeriesSpacing(chart, slot, barWidth);
                    var seriesOffset = (chart.Series.Count - 1 - s - 1f) * seriesSpacing;
                    var categoryRatio = ResolveClusteredCategoryRatio(
                        chart, i, categoryIndex, pointCount, categorySlotCount,
                        seriesSpacing, seriesOffset);
                    var left = Interpolate(topLeft, bottomLeft, categoryRatio);
                    var right = Interpolate(topRight, bottomRight, categoryRatio);
                    ShiftBarToWallMidpoint(
                        left,
                        Interpolate(wallFrontTopLeft, wallFrontBottomLeft, categoryRatio),
                        ref left,
                        ref right);
                    var end = Interpolate(left, right, ProjectValueRatio((float)ratio));
                    DrawBarSegment(canvas, left, end, thicknessDirection,
                        series.Color, barWidth, true);
                }
            }
        }

        private void DrawStackedBars(SKCanvas canvas,
            SKPoint topLeft, SKPoint topRight, SKPoint bottomLeft, SKPoint bottomRight,
            SKPoint wallFrontTopLeft, SKPoint wallFrontBottomLeft,
            SKPoint thicknessDirection, ParsedChart chart, double min, double max, float slot)
        {
            var pointCount = chart.Categories.Count;
            var widthFactor = UsesPerspectiveFan(chart) ? 0.42f : 0.5f;
            var barWidth = Math.Max(3f, slot * widthFactor);
            for (var i = 0; i < pointCount; i++)
            {
                var categoryIndex = pointCount - 1 - i;
                var categoryRatio = ResolveCategoryRatio(chart, i, categoryIndex, pointCount);
                var left = Interpolate(topLeft, bottomLeft, categoryRatio);
                var right = Interpolate(topRight, bottomRight, categoryRatio);
                ShiftBarToWallMidpoint(
                    left,
                    Interpolate(wallFrontTopLeft, wallFrontBottomLeft, categoryRatio),
                    ref left,
                    ref right);
                var positive = 0d;
                var negative = 0d;
                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    if (i >= series.Values.Count || !series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var value = NormalizeStackedValue(chart, i, series.Values[i].Value);
                    var startValue = value >= 0d ? positive : negative;
                    var endValue = startValue + value;
                    if (value >= 0d)
                    {
                        positive = endValue;
                    }
                    else
                    {
                        negative = endValue;
                    }

                    var startRatio = ProjectValueRatio((float)((startValue - min) / (max - min)));
                    var endRatio = ProjectValueRatio((float)((endValue - min) / (max - min)));
                    DrawBarSegment(
                        canvas,
                        Interpolate(left, right, startRatio),
                        Interpolate(left, right, endRatio),
                        thicknessDirection,
                        series.Color,
                        barWidth,
                        IsLastStackedSegment(chart, i, s, value));
                }
            }
        }

        private static double NormalizeStackedValue(ParsedChart chart, int categoryIndex, double value)
        {
            if (!chart.IsPercentStacked)
            {
                return value;
            }

            var total = 0d;
            var positive = value >= 0d;
            for (var i = 0; i < chart.Series.Count; i++)
            {
                var series = chart.Series[i];
                if (categoryIndex >= series.Values.Count || !series.Values[categoryIndex].HasValue)
                {
                    continue;
                }

                var candidate = series.Values[categoryIndex].Value;
                if ((candidate >= 0d) != positive)
                {
                    continue;
                }

                total += Math.Abs(candidate);
            }

            if (total <= 0d)
            {
                return 0d;
            }

            return value / total;
        }

        private static void ShiftBarToWallMidpoint(
            SKPoint wallBack, SKPoint wallFront, ref SKPoint left, ref SKPoint right)
        {
            var offset = new SKPoint(
                (wallFront.X - wallBack.X) * 0.5f,
                (wallFront.Y - wallBack.Y) * 0.5f);
            left = Offset(left, offset);
            right = Offset(right, offset);
        }

        private static bool IsLastStackedSegment(ParsedChart chart, int categoryIndex, int seriesIndex, double value)
        {
            var positive = value >= 0d;
            for (var i = seriesIndex + 1; i < chart.Series.Count; i++)
            {
                var series = chart.Series[i];
                if (categoryIndex >= series.Values.Count || !series.Values[categoryIndex].HasValue)
                {
                    continue;
                }

                var laterValue = series.Values[categoryIndex].Value;
                if ((laterValue >= 0d) == positive && laterValue != 0d)
                {
                    return false;
                }
            }

            return true;
        }

        private int ResolveCategorySlotCount(ParsedChart chart, int pointCount)
        {
            if (!chart.HasDateCategoryAxis)
            {
                return pointCount;
            }

            double minDate;
            double maxDate;
            if (!ChartDateAxisMath.TryGetDateAxisRange(chart, out minDate, out maxDate))
            {
                return pointCount;
            }

            var min = DateSerialConverter.FromSerial(minDate, _dateSystem);
            var max = DateSerialConverter.FromSerial(maxDate, _dateSystem);
            var monthCount = (max.Year - min.Year) * 12 + max.Month - min.Month + 1;
            return Math.Max(pointCount, monthCount);
        }

        private float ResolveCategoryRatio(
            ParsedChart chart, int sourceIndex, int reversedIndex, int pointCount)
        {
            if (chart.HasDateCategoryAxis
                && sourceIndex < chart.CategoryValues.Count
                && chart.CategoryValues[sourceIndex].HasValue)
            {
                double plotMinDate;
                double plotMaxDate;
                if (ChartDateAxisMath.TryGetDateAxisPlotRange(
                    chart, _dateSystem, out plotMinDate, out plotMaxDate))
                {
                    var serial = ChartDateAxisMath.NormalizePointSerial(
                        chart, chart.CategoryValues[sourceIndex].Value, _dateSystem);
                    var ratio = (serial - plotMinDate) / (plotMaxDate - plotMinDate);
                    ratio = Math.Max(0d, Math.Min(1d, ratio));
                    return chart.CategoryAxisMinMax ? 1f - (float)ratio : (float)ratio;
                }
            }

            return (reversedIndex + 0.5f) / pointCount;
        }

        private float ResolveClusteredCategoryRatio(
            ParsedChart chart,
            int sourceIndex,
            int reversedIndex,
            int pointCount,
            int categorySlotCount,
            float seriesSpacing,
            float seriesOffset)
        {
            var baseRatio = ResolveCategoryRatio(chart, sourceIndex, reversedIndex, pointCount);
            if (chart.HasDateCategoryAxis)
            {
                var minimumBaseRatio = (0.6f + seriesSpacing) / categorySlotCount;
                if (baseRatio < minimumBaseRatio)
                {
                    baseRatio = minimumBaseRatio;
                }
            }

            return baseRatio + seriesOffset / categorySlotCount;
        }

        private static float ResolveClusteredSeriesSpacing(
            ParsedChart chart, float slot, float barWidth)
        {
            if (!chart.HasDateCategoryAxis)
            {
                return 0.20f;
            }

            return barWidth / Math.Max(0.01f, slot);
        }

        private static float ResolveClusteredBarWidth(ParsedChart chart, float slot)
        {
            if (chart.HasDateCategoryAxis)
            {
                return Math.Max(1f, slot * 0.25f);
            }

            return Math.Max(2.5f, slot * 0.25f);
        }

        private static void DrawBarSegment(SKCanvas canvas, SKPoint start, SKPoint end,
            SKPoint thicknessDirection, SKColor color, float width, bool drawEndFace)
        {
            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var length = (float)Math.Sqrt(dx * dx + dy * dy);
            if (length <= 0f)
            {
                return;
            }

            var halfWidth = width * 0.5f;
            var halfThickness = new SKPoint(
                thicknessDirection.X * halfWidth,
                thicknessDirection.Y * halfWidth);
            var frontTopStart = new SKPoint(
                start.X - halfThickness.X,
                start.Y - halfThickness.Y);
            var frontTopEnd = new SKPoint(
                end.X - halfThickness.X,
                end.Y - halfThickness.Y);
            var frontBottomStart = Offset(start, halfThickness);
            var frontBottomEnd = Offset(end, halfThickness);
            var depth = Math.Min(2f, Math.Max(0.75f, width * 0.20f));
            var depthOffset = new SKPoint(depth, -depth * 0.55f);
            var backTopStart = Offset(frontTopStart, depthOffset);
            var backTopEnd = Offset(frontTopEnd, depthOffset);

            DrawBarFace(canvas, Shade(color, 0.76f),
                frontTopStart, frontTopEnd, backTopEnd, backTopStart);
            DrawBarFace(canvas, color,
                frontTopStart, frontTopEnd, frontBottomEnd, frontBottomStart);
            if (drawEndFace)
            {
                var backBottomEnd = Offset(frontBottomEnd, depthOffset);
                DrawBarFace(canvas, Shade(color, 0.64f),
                    frontTopEnd, backTopEnd, backBottomEnd, frontBottomEnd);
            }
        }

        private static void DrawBarFace(SKCanvas canvas, SKColor color,
            SKPoint first, SKPoint second, SKPoint third, SKPoint fourth)
        {
            using (var path = new SKPath())
            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Fill,
                Color = color,
                IsAntialias = true
            })
            {
                path.MoveTo(first);
                path.LineTo(second);
                path.LineTo(third);
                path.LineTo(fourth);
                path.Close();
                canvas.DrawPath(path, paint);
            }
        }

        private static SKPoint Offset(SKPoint point, SKPoint offset)
        {
            return new SKPoint(point.X + offset.X, point.Y + offset.Y);
        }

        private static SKColor Shade(SKColor color, float factor)
        {
            return new SKColor(
                (byte)Math.Round(color.Red * factor),
                (byte)Math.Round(color.Green * factor),
                (byte)Math.Round(color.Blue * factor),
                color.Alpha);
        }

        private void DrawCategoryLabels(SKCanvas canvas, SKRect area,
            SKPoint topLeft, SKPoint bottomLeft, ParsedChart chart)
        {
            if (chart.HasDateCategoryAxis && DrawDateCategoryLabels(canvas, area, topLeft, bottomLeft, chart))
            {
                return;
            }

            var count = chart.Categories.Count;
            var labelStep = chart.HasMinorValueGridlines ? 1 : 2;
            for (var i = 0; i < count; i += labelStep)
            {
                var categoryIndex = count - 1 - i;
                var point = Interpolate(topLeft, bottomLeft, (categoryIndex + 0.5f) / count);
                DrawText(canvas, area, chart.Categories[i], point.X - 8f, point.Y, chart.AxisLabelColor, true);
            }
        }

        private bool DrawDateCategoryLabels(SKCanvas canvas, SKRect area,
            SKPoint topLeft, SKPoint bottomLeft, ParsedChart chart)
        {
            double minDate;
            double maxDate;
            double plotMinDate;
            double plotMaxDate;
            if (!ChartDateAxisMath.TryGetDateAxisRange(chart, out minDate, out maxDate)
                || !ChartDateAxisMath.TryGetDateAxisPlotRange(
                    chart, _dateSystem, out plotMinDate, out plotMaxDate))
            {
                return false;
            }

            var slotCount = ResolveCategorySlotCount(chart, chart.Categories.Count);
            var monthStep = chart.PerspectiveDeg >= 100d ? 4 : (slotCount > 8 ? 3 : 1);
            var current = ChartDateAxisMath.StartOfMonth(minDate, _dateSystem);
            var last = ChartDateAxisMath.StartOfMonth(maxDate, _dateSystem);
            while (current <= last + 0.5d)
            {
                var ratio = (current - plotMinDate) / (plotMaxDate - plotMinDate);
                ratio = Math.Max(0d, Math.Min(1d, ratio));
                var categoryRatio = chart.CategoryAxisMinMax ? 1f - (float)ratio : (float)ratio;
                var point = Interpolate(topLeft, bottomLeft, categoryRatio);
                var formatCode = string.IsNullOrEmpty(chart.CategoryFormatCode)
                    ? "mmm-yy"
                    : chart.CategoryFormatCode;
                var label = ChartXmlParser.FormatLabel(
                    current, formatCode, _dateSystem, _culture);
                DrawText(canvas, area, label != null ? label.TrimEnd() : string.Empty,
                    point.X - 8f, point.Y, chart.AxisLabelColor, true);
                current = ChartDateAxisMath.AddMonthsSerial(current, monthStep, _dateSystem);
            }

            return true;
        }

        private static SKPoint Interpolate(SKPoint start, SKPoint end, float ratio)
        {
            return new SKPoint(
                start.X + (end.X - start.X) * ratio,
                start.Y + (end.Y - start.Y) * ratio);
        }

        private static float ProjectValueRatio(float ratio)
        {
            return ratio / (1f + 0.20f * (1f - ratio));
        }

        private static float Distance(SKPoint first, SKPoint second)
        {
            var dx = second.X - first.X;
            var dy = second.Y - first.Y;
            return (float)Math.Sqrt(dx * dx + dy * dy);
        }

        private static SKPoint ResolveBarThicknessDirection(SKPoint wallTopLeft, SKPoint wallBottomLeft)
        {
            var length = Distance(wallTopLeft, wallBottomLeft);
            return new SKPoint(
                (wallBottomLeft.X - wallTopLeft.X) / length,
                (wallBottomLeft.Y - wallTopLeft.Y) / length);
        }

        private void DrawText(SKCanvas canvas, SKRect clip, string text, float x, float y, SKColor color, bool alignRight)
        {
            if (string.IsNullOrEmpty(text)) return;
            var fontValue = new FontValue { Name = "Calibri", Size = LabelSizePt };
            var fontContext = _context.GetFontContext(fontValue);
            var drawX = alignRight ? x - fontContext.Measure(text) : x;
            var runs = fontContext.SplitRuns(text);
            canvas.Save();
            canvas.ClipRect(clip);
            for (var i = 0; i < runs.Count; i++)
            {
                using (var font = new SKFont(runs[i].Typeface, LabelSizePt))
                {
                    PdfTextPathRenderer.DrawText(canvas, runs[i].Text, drawX, y + LabelSizePt * 0.35f, font, color);
                }
                drawX += runs[i].WidthPt;
            }
            canvas.Restore();
        }

        private void DrawCenteredText(SKCanvas canvas, SKRect clip, string text, float x, float y, SKColor color)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var fontValue = new FontValue { Name = "Calibri", Size = LabelSizePt };
            var fontContext = _context.GetFontContext(fontValue);
            DrawText(canvas, clip, text, x + fontContext.Measure(text) * 0.5f, y, color, true);
        }

        private static double NiceNumber(double value)
        {
            if (value <= 0d) return 1d;
            var exponent = Math.Floor(Math.Log10(value));
            var fraction = value / Math.Pow(10d, exponent);
            var nice = fraction < 1.5d ? 1d : (fraction < 3d ? 2d : (fraction < 7d ? 5d : 10d));
            return nice * Math.Pow(10d, exponent);
        }
    }
}
