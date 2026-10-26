using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Renders area3DChart series as depth-separated vertical area planes with a projected plot box.
    /// </summary>
    internal sealed class Area3DChartRenderer
    {
        private static readonly SKColor GridColor = new SKColor(0xD9, 0xD9, 0xD9);
        private const float LabelSizePt = 9f;
        private const float MajorGridlineWidthPt = 0.5f;
        private const float ValueAxisWidthPt = 0.5f;
        private const double FrontAxisGap = 0.06d;
        private const double StackedFrontAxisGap = 0.28d;

        private readonly RenderContext _context;
        private readonly DateSystem _dateSystem;
        private readonly CultureInfo _culture;

        public Area3DChartRenderer(RenderContext context, DateSystem dateSystem, CultureInfo culture)
        {
            _context = context;
            _dateSystem = dateSystem;
            _culture = culture ?? CultureInfo.CurrentCulture;
        }

        public void Draw(SKCanvas canvas, SKRect area, SKRect plot, ParsedChart chart, double min, double max, double major)
        {
            if (canvas == null || chart == null || chart.Series == null || chart.Series.Count == 0 || max <= min)
            {
                return;
            }

            var renderedMajor = ResolveMajorUnit(chart, min, max, major);
            DrawPlotBox(canvas, plot, chart, min, max, renderedMajor, _dateSystem);
            DrawSeries(canvas, plot, chart, min, max, _dateSystem);
            DrawValueLabels(canvas, area, plot, chart, min, max, renderedMajor);
            DrawCategoryLabels(canvas, area, plot, chart);
            DrawSeriesAxisLabels(canvas, area, plot, chart);
        }

        private static void DrawPlotBox(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, double major, DateSystem dateSystem)
        {
            var frontLeft = Project(plot, chart, 0d, 0d, 0d);
            var frontRight = Project(plot, chart, 1d, 0d, 0d);
            var frontTopLeft = Project(plot, chart, 0d, 1d, 0d);
            var backLeft = Project(plot, chart, 0d, 0d, 1d);
            var backRight = Project(plot, chart, 1d, 0d, 1d);
            var backTopLeft = Project(plot, chart, 0d, 1d, 1d);
            var backTopRight = Project(plot, chart, 1d, 1d, 1d);

            var wallColor = chart.HasChartAreaGradient ? SKColors.Transparent : SKColors.White;
            var floorColor = chart.HasChartAreaGradient ? SKColors.Transparent : new SKColor(0xF8, 0xF8, 0xF8);
            using (var floorFill = new SKPaint { Style = SKPaintStyle.Fill, Color = floorColor, IsAntialias = true })
            using (var border = new SKPaint { Style = SKPaintStyle.Stroke, Color = GridColor, StrokeWidth = MajorGridlineWidthPt, IsAntialias = true })
            {
                Area3DWallFillRenderer.Draw(canvas, frontLeft, frontTopLeft, backLeft, backTopLeft,
                    backRight, backTopRight, chart.SideWallImageData, chart.BackWallImageData, wallColor);
                canvas.DrawLine(backTopLeft, backTopRight, border);
                canvas.DrawLine(backLeft, backRight, border);

                using (var floor = new SKPath())
                {
                    floor.MoveTo(backLeft);
                    floor.LineTo(backRight);
                    floor.LineTo(frontRight);
                    floor.LineTo(frontLeft);
                    floor.Close();
                    canvas.DrawPath(floor, floorFill);
                    if (chart.IsStacked || chart.IsPercentStacked)
                    {
                        canvas.DrawLine(backLeft, backRight, border);
                        canvas.DrawLine(frontLeft, frontRight, border);
                        canvas.DrawLine(frontLeft, backLeft, border);
                    }
                    else
                    {
                        canvas.DrawPath(floor, border);
                    }
                }

                if (chart.HasMinorValueGridlines)
                {
                    var minor = chart.MinorUnit.HasValue && chart.MinorUnit.Value > 0d
                        ? chart.MinorUnit.Value
                        : major / 5d;
                    using (var minorGrid = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        Color = chart.MinorValueGridlineColor,
                        StrokeWidth = (float)Math.Max(0.5d, chart.MinorValueGridlineWidthPt),
                        IsAntialias = true
                    })
                    {
                        var epsilon = minor * 0.1d;
                        for (var value = min + minor; value < max - epsilon; value += minor)
                        {
                            var majorRatio = (value - min) / major;
                            if (Math.Abs(majorRatio - Math.Round(majorRatio)) <= 0.001d)
                            {
                                continue;
                            }

                            DrawProjectedValueGridline(canvas, plot, chart, (value - min) / (max - min), minorGrid);
                        }
                    }
                }

                var steps = (int)Math.Round((max - min) / major);
                for (var i = 0; i <= steps; i++)
                {
                    var value = min + i * major;
                    var normalized = (value - min) / (max - min);
                    DrawProjectedValueGridline(canvas, plot, chart, normalized, border);
                }

                var monthlyTicks = Area3DDateAxisMath.BuildMonthlyTicks(chart, dateSystem, chart.Culture);
                var categoryCount = monthlyTicks.Count > 0 ? monthlyTicks.Count : ResolvePointCount(chart);
                for (var i = 0; i < categoryCount; i++)
                {
                    var x = monthlyTicks.Count > 0 ? monthlyTicks[i].Item1 : Area3DDateAxisMath.CategoryPosition(chart, i, categoryCount, dateSystem);
                    var floorFront = Project(plot, chart, x, 0d, 0d);
                    var floorBack = Project(plot, chart, x, 0d, 1d);
                    if (!chart.IsStacked && !chart.IsPercentStacked)
                    {
                        canvas.DrawLine(floorFront, floorBack, border);
                    }
                    else
                    {
                        DrawCategoryAxisTick(canvas, floorFront, border);
                    }
                }

                DrawValueAxisLine(canvas, plot, chart, min, max, major);
            }
        }

        private static void DrawCategoryAxisTick(SKCanvas canvas, SKPoint axisPoint, SKPaint paint)
        {
            var end = new SKPoint(axisPoint.X, axisPoint.Y + 2.5f);
            canvas.DrawLine(axisPoint, end, paint);
        }

        private static void DrawValueAxisLine(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, double major)
        {
            if (!chart.HasValueAxisLine)
            {
                return;
            }

            using (var axis = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.ValueAxisLineColor,
                StrokeWidth = ValueAxisWidthPt,
                IsAntialias = true
            })
            {
                var axisX = chart.ValueAxisOnRight ? 1d : 0d;
                var axisDepth = chart.ValueAxisOnRight ? 1d : 0d;
                var bottom = Project(plot, chart, axisX, 0d, axisDepth);
                var top = Project(plot, chart, axisX, 1d, axisDepth);
                canvas.DrawLine(bottom, top, axis);

                if (string.Equals(chart.ValueAxisMajorTickMark, "none", StringComparison.Ordinal))
                {
                    return;
                }

                var steps = (int)Math.Round((max - min) / major);
                for (var i = 0; i <= steps; i++)
                {
                    var point = Project(plot, chart, axisX, (i * major) / (max - min), axisDepth);
                    var outward = chart.ValueAxisOnRight ? 1f : -1f;
                    var tickStart = point.X;
                    var tickEnd = point.X + 3f * outward;
                    if (string.Equals(chart.ValueAxisMajorTickMark, "cross", StringComparison.Ordinal))
                    {
                        tickStart = point.X - 3f;
                        tickEnd = point.X + 3f;
                    }
                    else if (string.Equals(chart.ValueAxisMajorTickMark, "in", StringComparison.Ordinal))
                    {
                        tickEnd = point.X - 3f * outward;
                    }

                    canvas.DrawLine(tickStart, point.Y, tickEnd, point.Y, axis);
                }
            }
        }

        private static void DrawProjectedValueGridline(SKCanvas canvas, SKRect plot, ParsedChart chart, double normalized, SKPaint paint)
        {
            var atFrontLeft = Project(plot, chart, 0d, normalized, 0d);
            var atBackLeft = Project(plot, chart, 0d, normalized, 1d);
            var atBackRight = Project(plot, chart, 1d, normalized, 1d);
            using (var line = new SKPath())
            {
                line.MoveTo(atFrontLeft);
                line.LineTo(atBackLeft);
                line.LineTo(atBackRight);
                canvas.DrawPath(line, paint);
            }
        }

        private static void DrawSeries(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, DateSystem dateSystem)
        {
            var pointCount = ResolvePointCount(chart);
            if (pointCount <= 0)
            {
                return;
            }

            if (chart.IsStacked || chart.IsPercentStacked)
            {
                DrawStackedSeries(canvas, plot, chart, min, max, pointCount, dateSystem);
                return;
            }

            for (var seriesIndex = chart.Series.Count - 1; seriesIndex >= 0; seriesIndex--)
            {
                var series = chart.Series[seriesIndex];
                var values = ChartStackingMath.BuildDisplayValues(chart, seriesIndex);
                var thickness = ResolveSeriesThickness(chart.Series.Count);
                var frontDepth = SeriesFrontDepth(seriesIndex, chart.Series.Count, thickness);
                var backDepth = frontDepth + thickness;
                var frontPoints = new List<SKPoint>();
                var backPoints = new List<SKPoint>();
                var xPositions = new List<double>();
                for (var i = 0; i < values.Count; i++)
                {
                    if (!values[i].HasValue)
                    {
                        continue;
                    }

                    var x = Area3DDateAxisMath.CategoryPosition(chart, i, pointCount, dateSystem);
                    var y = NormalizeValue(values[i].Value, min, max);
                    xPositions.Add(x);
                    frontPoints.Add(Project(plot, chart, x, y, frontDepth));
                    backPoints.Add(Project(plot, chart, x, y, backDepth));
                }

                if (frontPoints.Count == 0)
                {
                    continue;
                }

                Area3DDateAxisMath.SortByPosition(xPositions, frontPoints, backPoints);
                var baseline = NormalizeValue(Math.Max(min, 0d), min, max);
                var firstX = xPositions[0];
                var lastX = xPositions[xPositions.Count - 1];
                var frontFirstBase = Project(plot, chart, firstX, baseline, frontDepth);
                var frontLastBase = Project(plot, chart, lastX, baseline, frontDepth);
                var backFirstBase = Project(plot, chart, firstX, baseline, backDepth);
                var backLastBase = Project(plot, chart, lastX, baseline, backDepth);

                DrawAreaFace(canvas, backPoints, backFirstBase, backLastBase, Shade(series.Color, 0.84d), Shade(series.Color, 0.65d));
                using (var topFill = new SKPaint { Style = SKPaintStyle.Fill, Color = Shade(series.Color, 0.92d), IsAntialias = true })
                using (var sideFill = new SKPaint { Style = SKPaintStyle.Fill, Color = Shade(series.Color, 0.70d), IsAntialias = true })
                {
                    for (var pointIndex = 0; pointIndex < frontPoints.Count - 1; pointIndex++)
                    {
                        DrawQuad(canvas, frontPoints[pointIndex], frontPoints[pointIndex + 1], backPoints[pointIndex + 1], backPoints[pointIndex], topFill);
                    }

                    DrawQuad(canvas, frontPoints[0], backPoints[0], backFirstBase, frontFirstBase, sideFill);
                    DrawQuad(canvas, frontPoints[frontPoints.Count - 1], backPoints[backPoints.Count - 1], backLastBase, frontLastBase, sideFill);
                }

                DrawAreaFace(canvas, frontPoints, frontFirstBase, frontLastBase, series.Color, Shade(series.Color, 0.72d));
            }
        }

        private static void DrawAreaFace(SKCanvas canvas, List<SKPoint> points, SKPoint firstBase, SKPoint lastBase, SKColor fillColor, SKColor edgeColor)
        {
            using (var path = new SKPath())
            {
                path.MoveTo(points[0]);
                for (var i = 1; i < points.Count; i++)
                {
                    path.LineTo(points[i]);
                }

                path.LineTo(lastBase);
                path.LineTo(firstBase);
                path.Close();
                using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = fillColor, IsAntialias = true })
                using (var edge = new SKPaint { Style = SKPaintStyle.Stroke, Color = edgeColor, StrokeWidth = 1f, IsAntialias = true, StrokeJoin = SKStrokeJoin.Round })
                {
                    canvas.DrawPath(path, fill);
                    canvas.DrawPath(path, edge);
                }
            }
        }

        private static void DrawStackedSeries(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, int pointCount, DateSystem dateSystem)
        {
            var thickness = ResolveStackedSeriesThickness(chart);
            var frontDepth = StackedFrontAxisGap;
            var backDepth = frontDepth + thickness;
            var baseline = NormalizeValue(Math.Max(min, 0d), min, max);
            var frontUppers = new List<List<SKPoint>>();
            var frontLowers = new List<List<SKPoint>>();
            var backUppers = new List<List<SKPoint>>();
            var backLowers = new List<List<SKPoint>>();

            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
            {
                var upperValues = ChartStackingMath.BuildDisplayValues(chart, seriesIndex);
                var lowerValues = seriesIndex > 0
                    ? ChartStackingMath.BuildDisplayValues(chart, seriesIndex - 1)
                    : null;
                var frontUpper = new List<SKPoint>();
                var frontLower = new List<SKPoint>();
                var backUpper = new List<SKPoint>();
                var backLower = new List<SKPoint>();

                for (var i = 0; i < upperValues.Count; i++)
                {
                    if (!upperValues[i].HasValue)
                    {
                        continue;
                    }

                    var lower = baseline;
                    if (lowerValues != null && i < lowerValues.Count && lowerValues[i].HasValue)
                    {
                        lower = NormalizeValue(lowerValues[i].Value, min, max);
                    }

                    var x = Area3DDateAxisMath.CategoryPosition(chart, i, pointCount, dateSystem);
                    var upper = NormalizeValue(upperValues[i].Value, min, max);
                    frontUpper.Add(ProjectStackedPoint(plot, chart, x, upper, frontDepth));
                    frontLower.Add(ProjectStackedPoint(plot, chart, x, lower, frontDepth));
                    backUpper.Add(ProjectStackedPoint(plot, chart, x, upper, backDepth));
                    backLower.Add(ProjectStackedPoint(plot, chart, x, lower, backDepth));
                }

                frontUppers.Add(frontUpper);
                frontLowers.Add(frontLower);
                backUppers.Add(backUpper);
                backLowers.Add(backLower);
            }

            for (var seriesIndex = chart.Series.Count - 1; seriesIndex >= 0; seriesIndex--)
            {
                if (backUppers[seriesIndex].Count == 0)
                {
                    continue;
                }

                var series = chart.Series[seriesIndex];
                DrawBandFace(canvas, backUppers[seriesIndex], backLowers[seriesIndex], series, Shade(series.Color, 0.84d), false);
            }

            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
            {
                var frontUpper = frontUppers[seriesIndex];
                if (frontUpper.Count == 0)
                {
                    continue;
                }

                var frontLower = frontLowers[seriesIndex];
                var backUpper = backUppers[seriesIndex];
                var backLower = backLowers[seriesIndex];
                var series = chart.Series[seriesIndex];
                using (var sideFill = new SKPaint { Style = SKPaintStyle.Fill, Color = Shade(series.Color, 0.70d), IsAntialias = true })
                {
                    DrawQuad(canvas, frontUpper[0], backUpper[0], backLower[0], frontLower[0], sideFill);
                    var last = frontUpper.Count - 1;
                    DrawQuad(canvas, frontUpper[last], backUpper[last], backLower[last], frontLower[last], sideFill);
                }
            }

            var topSeriesIndex = chart.Series.Count - 1;
            if (topSeriesIndex >= 0 && frontUppers[topSeriesIndex].Count > 0)
            {
                var frontUpper = frontUppers[topSeriesIndex];
                var backUpper = backUppers[topSeriesIndex];
                var series = chart.Series[topSeriesIndex];
                var topShade = chart.IsPercentStacked ? 0.78d : 0.92d;
                using (var topFill = new SKPaint { Style = SKPaintStyle.Fill, Color = Shade(series.Color, topShade), IsAntialias = true })
                {
                    for (var pointIndex = 0; pointIndex < frontUpper.Count - 1; pointIndex++)
                    {
                        DrawQuad(canvas, frontUpper[pointIndex], frontUpper[pointIndex + 1], backUpper[pointIndex + 1], backUpper[pointIndex], topFill);
                    }
                }
            }

            for (var seriesIndex = chart.Series.Count - 1; seriesIndex >= 0; seriesIndex--)
            {
                if (frontUppers[seriesIndex].Count == 0)
                {
                    continue;
                }

                var series = chart.Series[seriesIndex];
                DrawBandFace(canvas, frontUppers[seriesIndex], frontLowers[seriesIndex], series, series.Color, true);
            }
        }

        private static SKPoint ProjectStackedPoint(SKRect plot, ParsedChart chart, double x, double y, double depth)
        {
            var point = Project(plot, chart, x, y, depth);
            if (!chart.IsPercentStacked)
            {
                return point;
            }

            if (Math.Abs(y - 1d) > 0.000001d)
            {
                return point;
            }

            var frontAnchor = Project(plot, chart, 0d, 1d, StackedFrontAxisGap);
            var currentLeft = Project(plot, chart, 0d, 1d, depth);
            var wallFrontLeft = Project(plot, chart, 0d, 1d, 0d);
            var wallBackLeft = Project(plot, chart, 0d, 1d, 1d);
            var wallBackRight = Project(plot, chart, 1d, 1d, 1d);
            var depthDx = wallBackLeft.X - wallFrontLeft.X;
            var backDx = wallBackRight.X - wallBackLeft.X;
            if (Math.Abs(depthDx) <= 0.000001f || Math.Abs(backDx) <= 0.000001f)
            {
                return point;
            }

            var depthSlope = (wallBackLeft.Y - wallFrontLeft.Y) / depthDx;
            var currentLeftY = frontAnchor.Y + depthSlope * (currentLeft.X - frontAnchor.X);
            var backSlope = (wallBackRight.Y - wallBackLeft.Y) / backDx;
            point.Y = currentLeftY + backSlope * (point.X - currentLeft.X);
            return point;
        }

        private static void DrawBandFace(SKCanvas canvas, List<SKPoint> upper, List<SKPoint> lower, ChartSeries series, SKColor fillColor, bool usePattern)
        {
            using (var path = new SKPath())
            {
                path.MoveTo(upper[0]);
                for (var i = 1; i < upper.Count; i++)
                {
                    path.LineTo(upper[i]);
                }

                for (var i = lower.Count - 1; i >= 0; i--)
                {
                    path.LineTo(lower[i]);
                }

                path.Close();
                if (usePattern && ChartPatternRenderer.DrawPath(canvas, path, series, fillColor))
                {
                    return;
                }

                using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = fillColor, IsAntialias = true })
                {
                    canvas.DrawPath(path, fill);
                }
            }
        }

        private static void DrawQuad(SKCanvas canvas, SKPoint first, SKPoint second, SKPoint third, SKPoint fourth, SKPaint paint)
        {
            using (var path = new SKPath())
            {
                path.MoveTo(first);
                path.LineTo(second);
                path.LineTo(third);
                path.LineTo(fourth);
                path.Close();
                canvas.DrawPath(path, paint);
            }
        }

        private void DrawValueLabels(SKCanvas canvas, SKRect area, SKRect plot, ParsedChart chart, double min, double max, double major)
        {
            var steps = (int)Math.Round((max - min) / major);
            for (var i = 0; i <= steps; i++)
            {
                var value = min + i * major;
                var normalized = (value - min) / (max - min);
                var axisX = chart.ValueAxisOnRight ? 1d : 0d;
                var axisDepth = chart.ValueAxisOnRight ? 1d : 0d;
                var point = Project(plot, chart, axisX, normalized, axisDepth);
                var label = ChartXmlParser.FormatLabel(value, chart.ValueFormatCode, _dateSystem, _culture);
                var labelX = point.X + (chart.ValueAxisOnRight ? 7f : -7f);
                DrawText(canvas, area, label != null ? label.TrimEnd() : string.Empty, labelX, point.Y, chart.AxisLabelColor, !chart.ValueAxisOnRight, false);
            }
        }

        private void DrawCategoryLabels(SKCanvas canvas, SKRect area, SKRect plot, ParsedChart chart)
        {
            var monthlyTicks = Area3DDateAxisMath.BuildMonthlyTicks(chart, _dateSystem, _culture);
            if (monthlyTicks.Count > 0)
            {
                for (var i = 0; i < monthlyTicks.Count; i++)
                {
                    var point = Project(plot, chart, monthlyTicks[i].Item1, 0d, 0d);
                    canvas.Save();
                    canvas.RotateDegrees((float)chart.CategoryAxisTextRotationDeg, point.X, point.Y);
                    DrawText(canvas, area, monthlyTicks[i].Item2, point.X - 4f, point.Y, chart.AxisLabelColor, true, false);
                    canvas.Restore();
                }

                return;
            }

            var count = Math.Max(chart.Categories.Count, chart.CategoryValues.Count);
            for (var i = 0; i < count; i++)
            {
                var position = Area3DDateAxisMath.CategoryPosition(chart, i, count, _dateSystem);
                var point = Project(plot, chart, position, 0d, 0d);
                var label = ResolveCategoryLabel(chart, i);
                DrawText(canvas, area, label, point.X, point.Y + 7f, chart.AxisLabelColor, false, true);
            }
        }

        private void DrawSeriesAxisLabels(SKCanvas canvas, SKRect area, SKRect plot, ParsedChart chart)
        {
            if (!chart.HasVisibleSeriesAxis)
            {
                return;
            }

            for (var i = chart.Series.Count - 1; i >= 0; i--)
            {
                var thickness = ResolveSeriesThickness(chart.Series.Count);
                var depth = SeriesFrontDepth(i, chart.Series.Count, thickness);
                var point = Project(plot, chart, 1d, 0d, depth);
                DrawText(canvas, area, chart.Series[i].Name, point.X + 5f, point.Y, chart.SeriesAxisLabelColor, false, false);
            }
        }

        private string ResolveCategoryLabel(ParsedChart chart, int index)
        {
            if (index < chart.Categories.Count && !string.IsNullOrEmpty(chart.Categories[index]))
            {
                return chart.Categories[index];
            }

            if (index < chart.CategoryValues.Count && chart.CategoryValues[index].HasValue)
            {
                var formatCode = string.IsNullOrEmpty(chart.CategoryFormatCode) ? "mmm-yy" : chart.CategoryFormatCode;
                var label = ChartXmlParser.FormatLabel(chart.CategoryValues[index].Value, formatCode, _dateSystem, _culture);
                return label != null ? label.TrimEnd() : string.Empty;
            }

            return string.Empty;
        }

        private void DrawText(SKCanvas canvas, SKRect clip, string text, float x, float y, SKColor color, bool alignRight, bool alignTop)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var fontValue = new FontValue { Name = "Calibri", Size = LabelSizePt };
            var fontContext = _context.GetFontContext(fontValue);
            using (var paint = new SKPaint
            {
                Color = color,
                IsAntialias = true,
                Typeface = _context.Fonts.Resolve(fontValue),
                TextSize = LabelSizePt,
                TextAlign = SKTextAlign.Left,
                SubpixelText = true
            })
            {
                var drawX = x;
                if (alignRight)
                {
                    drawX -= fontContext.Measure(text);
                }
                else if (alignTop)
                {
                    drawX -= fontContext.Measure(text) * 0.5f;
                }

                var metrics = paint.FontMetrics;
                var baseline = alignTop ? y - metrics.Ascent : y - (metrics.Ascent + metrics.Descent) * 0.5f;
                canvas.Save();
                canvas.ClipRect(clip);
                var runs = fontContext.SplitRuns(text);
                var cursor = drawX;
                for (var i = 0; i < runs.Count; i++)
                {
                    var run = runs[i];
                    using (var font = new SKFont(run.Typeface, LabelSizePt))
                    {
                        font.Subpixel = true;
                        PdfTextPathRenderer.DrawText(canvas, run.Text, cursor, baseline, font, color);
                    }

                    cursor += run.WidthPt;
                }

                canvas.Restore();
            }
        }

        private static SKPoint Project(SKRect plot, ParsedChart chart, double x, double y, double depth)
        {
            var depthScale = Clamp(chart.DepthPercent / 100d, 0.5d, 2d);
            var rotationScale = Clamp(Math.Abs(chart.RotationYDeg) / 20d, 0.5d, 1.5d);
            var depthBlend = Clamp(depth * depthScale * rotationScale, 0d, 1d);
            var tiltScale = Clamp(Math.Abs(chart.RotationXDeg) / 15d, 0.35d, 1.5d);
            var perspectiveScale = chart.RightAngleAxes
                ? 0.85d
                : Clamp(chart.PerspectiveDeg / 30d, 0.5d, 1.5d);
            var pointCount = ResolvePointCount(chart);
            var compact = chart.IsStacked || chart.IsPercentStacked
                ? 0d
                : Clamp((chart.Series.Count - pointCount) / 3d, 0d, 1d);

            var frontLeftX = plot.Left + plot.Width * Lerp(0.03f, 0.23f, compact);
            var frontRightX = plot.Left + plot.Width * Lerp(0.86f, 0.68f, compact);
            var backLeftX = frontLeftX + plot.Width * Lerp(0.19f, 0.29f, compact) * (float)perspectiveScale;
            var backRightX = frontRightX + plot.Width * Lerp(0.05f, 0.13f, compact) * (float)perspectiveScale;
            var frontBaseLeftY = plot.Top + plot.Height * Lerp(0.62f, 0.85f, compact);
            var frontBaseRightY = frontBaseLeftY + plot.Height * Lerp(0.26f, 0.16f, compact) * (float)tiltScale;
            var backBaseLeftY = frontBaseLeftY - plot.Height * Lerp(0.13f, 0.50f, compact) * (float)tiltScale;
            var backBaseRightY = backBaseLeftY + plot.Height * Lerp(0.09f, 0.10f, compact) * (float)tiltScale;
            var frontHeight = plot.Height * Lerp(0.56f, 0.72f, compact);
            var backHeight = plot.Height * Lerp(0.51f, 0.42f, compact);

            var leftX = Lerp(frontLeftX, backLeftX, depthBlend);
            var rightX = Lerp(frontRightX, backRightX, depthBlend);
            var baseLeftY = Lerp(frontBaseLeftY, backBaseLeftY, depthBlend);
            var baseRightY = Lerp(frontBaseRightY, backBaseRightY, depthBlend);
            var height = Lerp(frontHeight, backHeight, depthBlend);
            var projectedX = Lerp(leftX, rightX, x);
            if (chart.IsPercentStacked)
            {
                var perspectiveInset = plot.Width * 0.02f * (float)(1d - Clamp(y, 0d, 1d));
                projectedX += Lerp(perspectiveInset, -perspectiveInset, (float)Clamp(x, 0d, 1d));
            }

            var baselineY = Lerp(baseLeftY, baseRightY, x);
            var rightHeight = chart.IsPercentStacked ? height * 1.10f : height;
            var projectedHeight = Lerp(height, rightHeight, (float)Clamp(x, 0d, 1d));
            return new SKPoint(projectedX, baselineY - (float)y * projectedHeight);
        }

        private static double SeriesDepth(int seriesIndex, int seriesCount)
        {
            if (seriesCount <= 1)
            {
                return 0.5d;
            }

            return (double)seriesIndex / (seriesCount - 1);
        }

        private static double ResolveSeriesThickness(int seriesCount)
        {
            if (seriesCount <= 0)
            {
                return 0.04d;
            }

            return Clamp(0.24d / seriesCount, 0.04d, 0.075d);
        }

        private static double ResolveStackedSeriesThickness(ParsedChart chart)
        {
            var depthPercent = chart != null ? chart.DepthPercent : 100d;
            if (chart != null && chart.IsPercentStacked)
            {
                return Clamp(0.14d * depthPercent / 100d, 0.08d, 0.22d);
            }

            return Clamp(0.22d * depthPercent / 100d, 0.12d, 0.32d);
        }

        private static double SeriesFrontDepth(int seriesIndex, int seriesCount, double thickness)
        {
            return FrontAxisGap + SeriesDepth(seriesIndex, seriesCount) * (1d - FrontAxisGap - thickness);
        }

        private static int ResolvePointCount(ParsedChart chart)
        {
            var count = Math.Max(chart.Categories.Count, chart.CategoryValues.Count);
            for (var i = 0; i < chart.Series.Count; i++)
            {
                if (chart.Series[i].Values.Count > count)
                {
                    count = chart.Series[i].Values.Count;
                }
            }

            return count;
        }

        private static double NormalizeValue(double value, double min, double max)
        {
            return Clamp((value - min) / (max - min), 0d, 1d);
        }

        private static double ResolveMajorUnit(ParsedChart chart, double min, double max, double major)
        {
            if (major <= 0d)
            {
                return (max - min) / 4d;
            }

            var pointCount = chart != null ? ResolvePointCount(chart) : 0;
            if (chart != null && chart.Series.Count > pointCount)
            {
                return major;
            }

            var steps = (max - min) / major;
            if (steps <= 5d)
            {
                return major;
            }

            return NiceNumber((max - min) / 4d);
        }

        private static double NiceNumber(double value)
        {
            if (value <= 0d)
            {
                return 1d;
            }

            var exponent = Math.Floor(Math.Log10(value));
            var fraction = value / Math.Pow(10d, exponent);
            double nice;
            if (fraction < 1.5d)
            {
                nice = 1d;
            }
            else if (fraction < 3d)
            {
                nice = 2d;
            }
            else if (fraction < 7d)
            {
                nice = 5d;
            }
            else
            {
                nice = 10d;
            }

            return nice * Math.Pow(10d, exponent);
        }

        private static SKColor Shade(SKColor color, double factor)
        {
            return new SKColor(
                (byte)Math.Round(color.Red * factor),
                (byte)Math.Round(color.Green * factor),
                (byte)Math.Round(color.Blue * factor),
                color.Alpha);
        }

        private static float Lerp(float start, float end, double amount)
        {
            return start + (end - start) * (float)amount;
        }

        private static double Clamp(double value, double min, double max)
        {
            if (value < min)
            {
                return min;
            }

            if (value > max)
            {
                return max;
            }

            return value;
        }
    }
}
