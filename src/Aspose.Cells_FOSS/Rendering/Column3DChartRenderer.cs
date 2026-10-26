using System;
using System.Globalization;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class Column3DChartRenderer
    {
        private static readonly SKColor GridColor = new SKColor(0xD9, 0xD9, 0xD9);
        private const float LabelSizePt = 9f;

        private readonly RenderContext _context;
        private readonly DateSystem _dateSystem;
        private readonly CultureInfo _culture;
        private SKRect _chartArea;
        public Column3DChartRenderer(RenderContext context, DateSystem dateSystem, CultureInfo culture)
        {
            _context = context;
            _dateSystem = dateSystem;
            _culture = culture ?? CultureInfo.CurrentCulture;
        }

        public void Draw(SKCanvas canvas, SKRect area, ParsedChart chart, double min, double max, double major)
        {
            if (canvas == null || chart == null || chart.Series.Count == 0 || max <= min)
            {
                return;
            }

            _chartArea = area;
            if (Column3DGeometry.UsesSeriesDepth(chart)
                && string.Equals(chart.BarShape, "box", StringComparison.Ordinal))
            {
                new Column3DSceneRenderer(_context, _dateSystem, _culture)
                    .Draw(canvas, area, chart, min, max, major);
                return;
            }
            var projectionArea = ResolveProjectionArea(area, chart);
            DrawPlotBox(canvas, projectionArea, chart, min, max, major);
            DrawColumns(canvas, projectionArea, chart, min, max);
            DrawDataLabels(canvas, projectionArea, chart, min, max);
            DrawValueLabels(canvas, projectionArea, chart, min, max, major);
            DrawCategoryLabels(canvas, projectionArea, chart);
        }

        private static void DrawPlotBox(SKCanvas canvas, SKRect area, ParsedChart chart,
            double min, double max, double major)
        {
            using (var grid = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = GridColor,
                StrokeWidth = 0.75f,
                IsAntialias = true
            })
            {
                var wallDepth = ResolveFloorDepth(area, chart);
                Column3DWallRenderer.Draw(canvas, area, chart, wallDepth);
                if (chart.HasMinorValueGridlines)
                {
                    var minor = chart.MinorUnit.HasValue && chart.MinorUnit.Value > 0d
                        ? chart.MinorUnit.Value
                        : major / 5d;
                    if (minor > 0d)
                    {
                        using (var minorGrid = new SKPaint
                        {
                            Style = SKPaintStyle.Stroke,
                            Color = chart.MinorValueGridlineColor,
                            StrokeWidth = (float)Math.Max(0.25d, chart.MinorValueGridlineWidthPt),
                            IsAntialias = true
                        })
                        {
                            var minorSteps = (int)Math.Floor((max - min) / minor + 0.0001d);
                            for (var index = 1; index < minorSteps; index++)
                            {
                                var value = min + index * minor;
                                var majorRatio = (value - min) / major;
                                if (Math.Abs(majorRatio - Math.Round(majorRatio)) < 0.0001d)
                                {
                                    continue;
                                }
                                var ratio = (value - min) / (max - min);
                                var left = Column3DGeometry.GridPoint(area, 0d, ratio);
                                var right = Column3DGeometry.GridPoint(area, 1d, ratio);
                                canvas.DrawLine(Offset(left, wallDepth), left, minorGrid);
                                canvas.DrawLine(left, right, minorGrid);
                            }
                        }
                    }
                }

                var steps = (int)Math.Round((max - min) / major);
                for (var index = 0; index <= steps; index++)
                {
                    var ratio = index * major / (max - min);
                    var left = Column3DGeometry.GridPoint(area, 0d, ratio);
                    var right = Column3DGeometry.GridPoint(area, 1d, ratio);
                    canvas.DrawLine(Offset(left, wallDepth), left, grid);
                    canvas.DrawLine(left, right, grid);
                }

                Column3DWallRenderer.DrawFloorGrid(canvas, area, chart, wallDepth,
                    grid, ResolvePointCount(chart));

                if (chart.HasValueAxisLine)
                {
                    using (var axis = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        Color = chart.ValueAxisLineColor,
                        StrokeWidth = (float)Math.Max(0.25d, chart.ValueAxisLineWidthPt),
                        IsAntialias = true
                    })
                    {
                        var axisBottom = ValueAxisPoint(area, 0d, chart);
                        var axisTop = ValueAxisPoint(area, 1d, chart);
                        canvas.DrawLine(axisBottom, axisTop, axis);
                        if (string.Equals(chart.ValueAxisMajorTickMark, "cross", StringComparison.Ordinal))
                        {
                            for (var index = 0; index <= steps; index++)
                            {
                                var ratio = index * major / (max - min);
                                var point = ValueAxisPoint(area, ratio, chart);
                                canvas.DrawLine(point.X - 3f, point.Y, point.X + 3f, point.Y, axis);
                            }
                        }
                    }
                }
            }
        }

        private static SKPoint ResolveFloorDepth(SKRect area, ParsedChart chart)
        {
            var defaultDepth = new SKPoint(-area.Width * 0.035f, area.Height * 0.014f);
            var pointCount = ResolvePointCount(chart);
            var barWidth = chart.IsStacked || chart.IsPercentStacked
                ? ResolveStackedColumnWidth(area, chart, pointCount)
                : area.Width * 0.022f;
            var shapeDepth = ResolveShapeDepth(barWidth, chart);
            var horizontalSlope = ResolveHorizontalAxisSlope(area, 0d);
            return Column3DGeometry.ResolveFloorDepth(area, chart, defaultDepth,
                barWidth, shapeDepth, horizontalSlope);
        }

        private static SKPoint ValueAxisPoint(SKRect area, double valueRatio, ParsedChart chart)
        {
            return Offset(Column3DGeometry.GridPoint(area, 0d, valueRatio),
                ResolveFloorDepth(area, chart));
        }

        private static void DrawColumns(SKCanvas canvas, SKRect area, ParsedChart chart, double min, double max)
        {
            if (chart.IsStacked || chart.IsPercentStacked)
            {
                DrawStackedColumns(canvas, area, chart, min, max);
                return;
            }

            DrawClusteredColumns(canvas, area, chart, min, max);
        }

        private static void DrawClusteredColumns(SKCanvas canvas, SKRect area, ParsedChart chart,
            double min, double max)
        {
            var pointCount = ResolvePointCount(chart);
            if (pointCount <= 0)
            {
                return;
            }

            var usesSeriesDepth = Column3DGeometry.UsesSeriesDepth(chart);
            var barWidth = area.Width * (usesSeriesDepth ? 0.035f : 0.022f);
            var seriesPitch = ResolveClusteredSeriesPitch(barWidth);
            var depth = ResolveShapeDepth(barWidth, chart);
            var groupSpan = (chart.Series.Count - 1) * seriesPitch;
            var horizontalAxisSlope = ResolveHorizontalAxisSlope(area, 0d);
            var verticalAxisSlope = ResolveVerticalAxisSlope(area, chart);
            var floorMidpointOffset = ResolveFloorMidpointOffset(area, chart);
            float heightExtension;
            Column3DGeometry.ResolveShapeOffsets(chart, barWidth, depth,
                horizontalAxisSlope, ref floorMidpointOffset, out heightExtension);

            for (var categoryIndex = 0; categoryIndex < pointCount; categoryIndex++)
            {
                var categoryRatio = (categoryIndex + 0.5d) / pointCount;
                var gridBaseline = Column3DGeometry.GridPoint(area, categoryRatio, 0d);
                var baseline = Offset(gridBaseline, floorMidpointOffset);
                var categoryDepth = usesSeriesDepth
                    ? Column3DGeometry.ResolveSeriesShapeDepth(area, barWidth, categoryRatio)
                    : depth;
                var firstSeries = usesSeriesDepth ? chart.Series.Count - 1 : 0;
                var lastSeries = usesSeriesDepth ? -1 : chart.Series.Count;
                var seriesStep = usesSeriesDepth ? -1 : 1;
                for (var seriesIndex = firstSeries; seriesIndex != lastSeries; seriesIndex += seriesStep)
                {
                    var series = chart.Series[seriesIndex];
                    if (categoryIndex >= series.Values.Count || !series.Values[categoryIndex].HasValue)
                    {
                        continue;
                    }

                    var valueRatio = (series.Values[categoryIndex].Value - min) / (max - min);
                    if (valueRatio < 0d)
                    {
                        valueRatio = 0d;
                    }
                    else if (valueRatio > 1d)
                    {
                        valueRatio = 1d;
                    }

                    var centerOffset = seriesIndex * seriesPitch - groupSpan * 0.5f;
                    var projectedOffset = ProjectAlongHorizontalAxis(centerOffset, horizontalAxisSlope);
                    var bottomCenter = usesSeriesDepth
                        ? Offset(gridBaseline, Column3DGeometry.ResolveSeriesFloorOffset(
                            Column3DGeometry.ResolveSeriesWallDepth(area, categoryRatio),
                            seriesIndex, chart.Series.Count))
                        : new SKPoint(baseline.X + projectedOffset.X,
                            baseline.Y + projectedOffset.Y);
                    var height = ValueHeight(area, categoryRatio, valueRatio)
                        + (float)valueRatio * heightExtension;
                    var heightOffset = ProjectAlongVerticalAxis(height, verticalAxisSlope);
                    DrawColumnShape(canvas, bottomCenter, barWidth, heightOffset, categoryDepth,
                        horizontalAxisSlope, usesSeriesDepth ? ResolveHorizontalAxisSlope(
                            area, valueRatio) : horizontalAxisSlope, chart, series, true);
                }
            }
        }

        private static void DrawStackedColumns(SKCanvas canvas, SKRect area, ParsedChart chart,
            double min, double max)
        {
            var pointCount = ResolvePointCount(chart);
            if (pointCount <= 0)
            {
                return;
            }

            var barWidth = ResolveStackedColumnWidth(area, chart, pointCount);
            var depth = ResolveShapeDepth(barWidth, chart);
            var horizontalAxisSlope = ResolveHorizontalAxisSlope(area, 0d);
            var verticalAxisSlope = ResolveVerticalAxisSlope(area, chart);
            var floorMidpointOffset = ResolveFloorMidpointOffset(area, chart);
            float heightExtension;
            Column3DGeometry.ResolveShapeOffsets(chart, barWidth, depth,
                horizontalAxisSlope, ref floorMidpointOffset, out heightExtension);

            for (var categoryIndex = 0; categoryIndex < pointCount; categoryIndex++)
            {
                var categoryRatio = (categoryIndex + 0.5d) / pointCount;
                var baseline = Offset(Column3DGeometry.GridPoint(area, categoryRatio, 0d),
                    floorMidpointOffset);
                for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
                {
                    var series = chart.Series[seriesIndex];
                    if (categoryIndex >= series.Values.Count || !series.Values[categoryIndex].HasValue)
                    {
                        continue;
                    }

                    double startValue;
                    double endValue;
                    ResolveStackedSegment(chart, categoryIndex, seriesIndex, out startValue, out endValue);
                    if (Math.Abs(endValue - startValue) < 0.0000001d)
                    {
                        continue;
                    }

                    var startRatio = ClampRatio((startValue - min) / (max - min));
                    var endRatio = ClampRatio((endValue - min) / (max - min));
                    var startHeight = ValueHeight(area, categoryRatio, startRatio)
                        + (float)startRatio * heightExtension;
                    var endHeight = ValueHeight(area, categoryRatio, endRatio)
                        + (float)endRatio * heightExtension;
                    var bottomCenter = Offset(
                        baseline,
                        ProjectAlongVerticalAxis(startHeight, verticalAxisSlope));
                    var heightOffset = ProjectAlongVerticalAxis(
                        endHeight - startHeight,
                        verticalAxisSlope);
                    var drawTopFace = IsLastStackedSegment(
                        chart,
                        categoryIndex,
                        seriesIndex,
                        endValue - startValue);
                    var bottomFaceSlope = Column3DGeometry.IsPercentStackedBox(chart)
                        ? ResolveHorizontalAxisSlope(area, startRatio) : horizontalAxisSlope;
                    var topFaceSlope = Column3DGeometry.IsPercentStackedBox(chart)
                        ? ResolveHorizontalAxisSlope(area, endRatio) : horizontalAxisSlope;
                    DrawColumnShape(canvas, bottomCenter, barWidth, heightOffset, depth,
                        bottomFaceSlope, topFaceSlope, chart, series, drawTopFace);
                }
            }
        }

        private static float ResolveStackedColumnWidth(SKRect area, ParsedChart chart, int pointCount)
        {
            if (pointCount <= 0)
            {
                return 0f;
            }

            var left = Column3DGeometry.GridPoint(area, 0d, 0d);
            var right = Column3DGeometry.GridPoint(area, 1d, 0d);
            var categoryPitch = (right.X - left.X) / pointCount;
            var gapRatio = Math.Max(0d, chart.GapWidthPercent / 100d);
            return categoryPitch / (float)(1d + gapRatio);
        }

        private static void ResolveStackedSegment(ParsedChart chart, int categoryIndex,
            int seriesIndex, out double startValue, out double endValue)
        {
            startValue = 0d;
            endValue = 0d;
            if (chart == null || seriesIndex < 0 || seriesIndex >= chart.Series.Count)
            {
                return;
            }

            var targetSeries = chart.Series[seriesIndex];
            if (categoryIndex < 0 || categoryIndex >= targetSeries.Values.Count
                || !targetSeries.Values[categoryIndex].HasValue)
            {
                return;
            }

            var targetValue = NormalizeStackedValue(
                chart,
                categoryIndex,
                targetSeries.Values[categoryIndex].Value);
            var positive = targetValue >= 0d;
            for (var index = 0; index < seriesIndex; index++)
            {
                var series = chart.Series[index];
                if (categoryIndex >= series.Values.Count || !series.Values[categoryIndex].HasValue)
                {
                    continue;
                }

                var value = NormalizeStackedValue(chart, categoryIndex, series.Values[categoryIndex].Value);
                if ((value >= 0d) == positive)
                {
                    startValue += value;
                }
            }

            endValue = startValue + targetValue;
        }

        private static double NormalizeStackedValue(ParsedChart chart, int categoryIndex, double value)
        {
            if (!chart.IsPercentStacked)
            {
                return value;
            }

            var total = 0d;
            var positive = value >= 0d;
            for (var index = 0; index < chart.Series.Count; index++)
            {
                var series = chart.Series[index];
                if (categoryIndex >= series.Values.Count || !series.Values[categoryIndex].HasValue)
                {
                    continue;
                }

                var candidate = series.Values[categoryIndex].Value;
                if ((candidate >= 0d) == positive)
                {
                    total += Math.Abs(candidate);
                }
            }

            return total > 0d ? value / total : 0d;
        }

        private static bool IsLastStackedSegment(ParsedChart chart, int categoryIndex,
            int seriesIndex, double value)
        {
            var positive = value >= 0d;
            for (var index = seriesIndex + 1; index < chart.Series.Count; index++)
            {
                var series = chart.Series[index];
                if (categoryIndex >= series.Values.Count || !series.Values[categoryIndex].HasValue)
                {
                    continue;
                }

                var laterValue = NormalizeStackedValue(
                    chart,
                    categoryIndex,
                    series.Values[categoryIndex].Value);
                if ((laterValue >= 0d) == positive && Math.Abs(laterValue) > 0.0000001d)
                {
                    return false;
                }
            }

            return true;
        }

        private static double ClampRatio(double value)
        {
            return Math.Max(0d, Math.Min(1d, value));
        }

        private static SKPoint ResolveFloorMidpointOffset(SKRect area, ParsedChart chart)
        {
            var floorDepth = ResolveFloorDepth(area, chart);
            return new SKPoint(floorDepth.X * 0.5f, floorDepth.Y * 0.5f);
        }

        private static float ResolveClusteredSeriesPitch(float barWidth)
        {
            return barWidth;
        }

        private static SKPoint ResolveColumnDepth(float barWidth)
        {
            return new SKPoint(barWidth * 0.55f, -barWidth * 0.24f);
        }

        private static SKPoint ResolveShapeDepth(float barWidth, ParsedChart chart)
        {
            if (Column3DGeometry.IsPercentStackedBox(chart))
            {
                return new SKPoint(barWidth * 0.55f, -barWidth * 0.275f);
            }
            return ResolveColumnDepth(barWidth);
        }

        private static SKRect ResolveProjectionArea(SKRect area, ParsedChart chart)
        {
            if (Column3DGeometry.UsesSeriesDepth(chart))
            {
                return new SKRect(
                    area.Left + area.Width * 0.13f,
                    area.Top + area.Height * 0.08f,
                    area.Right - area.Width * 0.04f,
                    area.Bottom - area.Height * 0.17f);
            }

            if (Column3DGeometry.IsPercentStackedBox(chart))
            {
                return new SKRect(area.Left, area.Top, area.Right,
                    area.Bottom - area.Height * 0.043f);
            }

            if (chart == null || !string.Equals(chart.BarShape, "cylinder", StringComparison.Ordinal))
            {
                return area;
            }

            return new SKRect(
                area.Left + area.Width * 0.12f,
                area.Top + area.Height * 0.10f,
                area.Right - area.Width * 0.18f,
                area.Bottom - area.Height * 0.20f);
        }

        private static float ResolveHorizontalAxisSlope(SKRect area, double valueRatio)
        {
            var left = Column3DGeometry.GridPoint(area, 0d, valueRatio);
            var right = Column3DGeometry.GridPoint(area, 1d, valueRatio);
            var width = right.X - left.X;
            if (Math.Abs(width) < 0.0001f)
            {
                return 0f;
            }
            return (right.Y - left.Y) / width;
        }

        private static SKPoint ProjectAlongHorizontalAxis(float distance, float slope)
        {
            return new SKPoint(distance, distance * slope);
        }

        private static float ResolveVerticalAxisSlope(SKRect area, ParsedChart chart)
        {
            var bottom = ValueAxisPoint(area, 0d, chart);
            var top = ValueAxisPoint(area, 1d, chart);
            var height = bottom.Y - top.Y;
            if (Math.Abs(height) < 0.0001f)
            {
                return 0f;
            }
            return (top.X - bottom.X) / height;
        }

        private static SKPoint ProjectAlongVerticalAxis(float height, float slope)
        {
            return new SKPoint(height * slope, -height);
        }

        private static void DrawColumnShape(SKCanvas canvas, SKPoint bottomCenter, float width,
            SKPoint heightOffset, SKPoint depth, float bottomAxisSlope, float topAxisSlope,
            ParsedChart chart, ChartSeries series, bool drawTopFace)
        {
            if (string.Equals(chart.BarShape, "cylinder", StringComparison.Ordinal))
            {
                Column3DCylinderRenderer.Draw(canvas, bottomCenter, width, heightOffset,
                    depth, bottomAxisSlope, series, drawTopFace);
                return;
            }

            DrawPrism(canvas, bottomCenter, width, heightOffset, depth,
                bottomAxisSlope, topAxisSlope, series, drawTopFace);
        }

        private static void DrawPrism(SKCanvas canvas, SKPoint bottomCenter, float width,
            SKPoint heightOffset, SKPoint depth, float bottomAxisSlope, float topAxisSlope,
            ChartSeries series, bool drawTopFace)
        {
            var color = series.Color;
            var halfWidth = width * 0.5f;
            var bottomSlopeOffset = width * bottomAxisSlope * 0.5f;
            var topSlopeOffset = width * topAxisSlope * 0.5f;
            var frontBottomLeft = new SKPoint(
                bottomCenter.X - halfWidth, bottomCenter.Y - bottomSlopeOffset);
            var frontBottomRight = new SKPoint(
                bottomCenter.X + halfWidth, bottomCenter.Y + bottomSlopeOffset);
            var topCenter = Offset(bottomCenter, heightOffset);
            var frontTopLeft = new SKPoint(topCenter.X - halfWidth, topCenter.Y - topSlopeOffset);
            var frontTopRight = new SKPoint(topCenter.X + halfWidth, topCenter.Y + topSlopeOffset);
            var backTopLeft = Offset(frontTopLeft, depth);
            var backTopRight = Offset(frontTopRight, depth);
            var backBottomRight = Offset(frontBottomRight, depth);

            using (var frontFill = new SKPaint { Style = SKPaintStyle.Fill, Color = color, IsAntialias = true })
            using (var topFill = new SKPaint { Style = SKPaintStyle.Fill, Color = Shade(color, 1.08d), IsAntialias = true })
            using (var sideFill = new SKPaint { Style = SKPaintStyle.Fill, Color = Shade(color, 0.72d), IsAntialias = true })
            using (var edge = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = series.LineColor,
                StrokeWidth = (float)Math.Max(0.25d, series.LineWidthPt),
                IsAntialias = true
            })
            {
                DrawQuad(canvas, frontTopRight, backTopRight, backBottomRight, frontBottomRight,
                    sideFill, edge, series.HasVisibleLine);
                if (drawTopFace)
                {
                    DrawQuad(canvas, frontTopLeft, backTopLeft, backTopRight, frontTopRight,
                        topFill, edge, series.HasVisibleLine);
                }
                using (var front = new SKPath())
                {
                    front.MoveTo(frontTopLeft);
                    front.LineTo(frontTopRight);
                    front.LineTo(frontBottomRight);
                    front.LineTo(frontBottomLeft);
                    front.Close();
                    canvas.DrawPath(front, frontFill);
                    if (series.HasVisibleLine)
                    {
                        canvas.DrawPath(front, edge);
                    }
                }
            }
        }

        private void DrawDataLabels(SKCanvas canvas, SKRect area, ParsedChart chart,
            double min, double max)
        {
            var pointCount = ResolvePointCount(chart);
            if (pointCount <= 0 || max <= min)
            {
                return;
            }

            var barWidth = area.Width * 0.022f;
            var seriesPitch = ResolveClusteredSeriesPitch(barWidth);
            var groupSpan = (chart.Series.Count - 1) * seriesPitch;
            var horizontalAxisSlope = ResolveHorizontalAxisSlope(area, 0d);
            var verticalAxisSlope = ResolveVerticalAxisSlope(area, chart);
            var floorMidpointOffset = ResolveFloorMidpointOffset(area, chart);

            for (var categoryIndex = 0; categoryIndex < pointCount; categoryIndex++)
            {
                var categoryRatio = (categoryIndex + 0.5d) / pointCount;
                var baseline = Offset(Column3DGeometry.GridPoint(area, categoryRatio, 0d),
                    floorMidpointOffset);
                for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
                {
                    var series = chart.Series[seriesIndex];
                    if (!series.ShowValueDataLabels || categoryIndex >= series.Values.Count
                        || !series.Values[categoryIndex].HasValue)
                    {
                        continue;
                    }

                    var value = series.Values[categoryIndex].Value;
                    var valueRatio = Math.Max(0d, Math.Min(1d, (value - min) / (max - min)));
                    var centerOffset = seriesIndex * seriesPitch - groupSpan * 0.5f;
                    var projectedOffset = ProjectAlongHorizontalAxis(centerOffset, horizontalAxisSlope);
                    var bottomCenter = new SKPoint(
                        baseline.X + projectedOffset.X,
                        baseline.Y + projectedOffset.Y);
                    var height = ValueHeight(area, categoryRatio, valueRatio);
                    var heightOffset = ProjectAlongVerticalAxis(height, verticalAxisSlope);
                    var labelX = bottomCenter.X + heightOffset.X;
                    var labelY = bottomCenter.Y + heightOffset.Y - 6f;
                    if (string.Equals(series.DataLabelPosition, "ctr", StringComparison.Ordinal))
                    {
                        labelX = bottomCenter.X + heightOffset.X * 0.5f;
                        labelY = bottomCenter.Y + heightOffset.Y * 0.5f;
                    }

                    var formatCode = !string.IsNullOrEmpty(series.ValueFormatCode)
                        ? series.ValueFormatCode
                        : chart.ValueFormatCode;
                    var label = ChartXmlParser.FormatLabel(value, formatCode, _dateSystem, _culture);
                    DrawText(canvas, area, label != null ? label.TrimEnd() : string.Empty,
                        labelX, labelY, series.DataLabelColor, false);
                }
            }
        }

        private static void DrawQuad(SKCanvas canvas, SKPoint a, SKPoint b, SKPoint c,
            SKPoint d, SKPaint fill, SKPaint edge, bool drawEdge)
        {
            using (var path = new SKPath())
            {
                path.MoveTo(a);
                path.LineTo(b);
                path.LineTo(c);
                path.LineTo(d);
                path.Close();
                canvas.DrawPath(path, fill);
                if (drawEdge)
                {
                    canvas.DrawPath(path, edge);
                }
            }
        }

        private void DrawValueLabels(SKCanvas canvas, SKRect area, ParsedChart chart,
            double min, double max, double major)
        {
            var steps = (int)Math.Round((max - min) / major);
            for (var index = 0; index <= steps; index++)
            {
                var value = min + index * major;
                var ratio = (value - min) / (max - min);
                var point = ValueAxisPoint(area, ratio, chart);
                var label = ChartXmlParser.FormatLabel(value, chart.ValueFormatCode,
                    _dateSystem, _culture);
                DrawText(canvas, area, label != null ? label.TrimEnd() : string.Empty,
                    point.X - 8f, point.Y, chart.AxisLabelColor, true);
            }
        }

        private void DrawCategoryLabels(SKCanvas canvas, SKRect area, ParsedChart chart)
        {
            var count = ResolvePointCount(chart);
            for (var index = 0; index < count; index++)
            {
                var ratio = (index + 0.5d) / count;
                var point = Column3DGeometry.UsesSeriesDepth(chart)
                    ? Column3DGeometry.ResolveSeriesFloorPoint(area, ratio, 1d)
                    : Offset(Column3DGeometry.GridPoint(area, ratio, 0d),
                        ResolveFloorDepth(area, chart));
                var label = index < chart.Categories.Count ? chart.Categories[index] : string.Empty;
                if (string.Equals(chart.BarShape, "cylinder", StringComparison.Ordinal))
                {
                    DrawRotatedText(canvas, label, point.X, point.Y + 15f,
                        chart.AxisLabelColor, -45f);
                }
                else
                {
                    DrawText(canvas, area, label, point.X, point.Y + 13f,
                        chart.AxisLabelColor, false);
                }
            }
        }

        private void DrawRotatedText(SKCanvas canvas, string text, float x, float y,
            SKColor color, float angleDegrees)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var fontValue = new FontValue { Name = "Calibri", Size = LabelSizePt };
            var fontContext = _context.GetFontContext(fontValue);
            var runs = fontContext.SplitRuns(text);
            if (runs.Count == 0)
            {
                return;
            }

            using (var metricsFont = new SKFont(runs[0].Typeface, LabelSizePt))
            {
                var metrics = metricsFont.Metrics;
                var baseline = -(metrics.Ascent + metrics.Descent) * 0.5f;
                var cursor = -fontContext.Measure(text) * 0.5f;
                canvas.Save();
                canvas.ClipRect(_chartArea);
                canvas.Translate(x, y);
                canvas.RotateDegrees(angleDegrees);
                for (var index = 0; index < runs.Count; index++)
                {
                    using (var font = new SKFont(runs[index].Typeface, LabelSizePt))
                    {
                        font.Subpixel = true;
                        PdfTextPathRenderer.DrawText(canvas, runs[index].Text,
                            cursor, baseline, font, color);
                    }
                    cursor += runs[index].WidthPt;
                }
                canvas.Restore();
            }
        }

        private void DrawText(SKCanvas canvas, SKRect clip, string text, float x, float y,
            SKColor color, bool alignRight)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var fontValue = new FontValue { Name = "Calibri", Size = LabelSizePt };
            var fontContext = _context.GetFontContext(fontValue);
            var drawX = alignRight ? x - fontContext.Measure(text) : x - fontContext.Measure(text) * 0.5f;
            var runs = fontContext.SplitRuns(text);
            if (runs.Count == 0)
            {
                return;
            }

            using (var metricsFont = new SKFont(runs[0].Typeface, LabelSizePt))
            {
                var metrics = metricsFont.Metrics;
                var baseline = y - (metrics.Ascent + metrics.Descent) * 0.5f;
                canvas.Save();
                var textClip = _chartArea.Width > 0f && _chartArea.Height > 0f
                    ? _chartArea
                    : clip;
                canvas.ClipRect(textClip);
                var cursor = drawX;
                for (var index = 0; index < runs.Count; index++)
                {
                    using (var font = new SKFont(runs[index].Typeface, LabelSizePt))
                    {
                        font.Subpixel = true;
                        PdfTextPathRenderer.DrawText(canvas, runs[index].Text, cursor, baseline, font, color);
                    }
                    cursor += runs[index].WidthPt;
                }
                canvas.Restore();
            }
        }

        private static float ValueHeight(SKRect area, double horizontalRatio, double valueRatio)
        {
            var bottom = Column3DGeometry.GridPoint(area, horizontalRatio, 0d);
            var top = Column3DGeometry.GridPoint(area, horizontalRatio, valueRatio);
            return bottom.Y - top.Y;
        }

        private static int ResolvePointCount(ParsedChart chart)
        {
            var count = Math.Max(chart.Categories.Count, chart.CategoryValues.Count);
            for (var index = 0; index < chart.Series.Count; index++)
            {
                if (chart.Series[index].Values.Count > count)
                {
                    count = chart.Series[index].Values.Count;
                }
            }
            return count;
        }

        private static SKPoint Interpolate(SKPoint start, SKPoint end, double ratio)
        {
            return new SKPoint(
                start.X + (end.X - start.X) * (float)ratio,
                start.Y + (end.Y - start.Y) * (float)ratio);
        }

        private static SKPoint Offset(SKPoint point, SKPoint offset)
        {
            return new SKPoint(point.X + offset.X, point.Y + offset.Y);
        }

        private static SKColor Shade(SKColor color, double factor)
        {
            var red = Math.Max(0d, Math.Min(255d, color.Red * factor));
            var green = Math.Max(0d, Math.Min(255d, color.Green * factor));
            var blue = Math.Max(0d, Math.Min(255d, color.Blue * factor));
            return new SKColor((byte)Math.Round(red), (byte)Math.Round(green), (byte)Math.Round(blue), color.Alpha);
        }
    }
}
