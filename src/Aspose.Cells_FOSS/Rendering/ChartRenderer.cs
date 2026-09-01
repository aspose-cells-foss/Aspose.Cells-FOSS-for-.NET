using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Draws a <see cref="ParsedChart"/> into a rectangle using SkiaSharp. Supports line, area,
    /// column/bar, radar, and pie plots.
    /// </summary>
    internal sealed class ChartRenderer
    {
        private static readonly SKColor GridColor = new SKColor(0xD9, 0xD9, 0xD9);
        private static readonly SKColor BorderColor = new SKColor(0xD9, 0xD9, 0xD9);
        private const float LabelSizePt = 9f;
        private const float StandardDateAxisGridlineRightExtensionPt = 8f;
        private const float StackedDateAxisGridlineRightExtensionPt = 10f;

        private readonly RenderContext _context;
        private readonly DateSystem _dateSystem;
        private CultureInfo _chartCulture;
        private bool _suppressCategoryAxisLabels;

        public ChartRenderer(RenderContext context)
        {
            _context = context;
            _dateSystem = context.Workbook != null && context.Workbook.Settings != null
                ? context.Workbook.Settings.DateSystem
                : DateSystem.Windows1900;
        }

        public void Draw(SKCanvas canvas, SKRect area, ParsedChart chart)
        {
            Draw(canvas, null, area, chart, false);
        }

        public void Draw(SKCanvas canvas, PageLayout page, SKRect area, ParsedChart chart)
        {
            Draw(canvas, page, area, chart, false);
        }

        public void Draw(SKCanvas canvas, PageLayout page, SKRect area, ParsedChart chart, bool suppressCategoryAxisLabels)
        {
            _chartCulture = chart != null && chart.Culture != null ? chart.Culture : _context.Culture;
            _suppressCategoryAxisLabels = suppressCategoryAxisLabels;

            DrawFrame(canvas, area, chart);

            if (chart.Kind == ChartKind.Pie)
            {
                DrawPie(canvas, page, area, chart);
                return;
            }

            if (chart.Series.Count == 0)
            {
                DrawTitle(canvas, page, area, chart, area.MidY);
                return;
            }

            var plot = ResolvePlotRect(area, chart);
            var seriesPlot = plot;

            double min, max, major;
            ComputeValueAxis(chart, out min, out max, out major);

            var titleCenterY = area.Top + (seriesPlot.Top - area.Top) * 0.5f;
            if (chart.Kind == ChartKind.Radar)
            {
                titleCenterY = area.Top + (seriesPlot.Top - area.Top) * 0.55f;
            }

            DrawTitle(canvas, page, area, chart, titleCenterY);
            if (chart.Kind == ChartKind.Radar)
            {
                DrawRadar(canvas, page, area, seriesPlot, chart, min, max, major);
                DrawPlotAreaBorder(canvas, page, area, seriesPlot, chart);
                DrawLegend(canvas, page, area, seriesPlot, chart);
                return;
            }

            if (chart.Kind == ChartKind.Bar)
            {
                DrawBarValueAxis(canvas, page, area, seriesPlot, chart, min, max, major);
                DrawSeries(canvas, seriesPlot, chart, min, max);
                DrawBarCategoryAxis(canvas, page, area, seriesPlot, chart);
                DrawPlotAreaBorder(canvas, page, area, seriesPlot, chart);
                DrawLegend(canvas, page, area, seriesPlot, chart);
                return;
            }
            else
            {
                DrawValueAxis(canvas, page, area, seriesPlot, chart, min, max, major);
                DrawCategoryAxis(canvas, page, area, seriesPlot, chart);
            }
            DrawSeries(canvas, seriesPlot, chart, min, max);
            DrawPlotAreaBorder(canvas, page, area, seriesPlot, chart);
            DrawLegend(canvas, page, area, seriesPlot, chart);
        }

        public void DrawFrame(SKCanvas canvas, SKRect area, ParsedChart chart)
        {
            if (canvas == null)
            {
                return;
            }

            var background = SKColors.White;
            if (chart != null && chart.HasChartAreaFill)
            {
                background = chart.ChartAreaFillColor;
            }

            using (var paint = new SKPaint { Style = SKPaintStyle.Fill, Color = background, IsAntialias = false })
            {
                canvas.DrawRect(area, paint);
            }

            using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, Color = BorderColor, StrokeWidth = 1f, IsAntialias = false })
            {
                canvas.DrawRect(area, paint);
            }
        }

        public SKRect GetVisualBounds(SKRect area, ParsedChart chart)
        {
            if (chart == null)
            {
                return area;
            }

            _chartCulture = chart.Culture != null ? chart.Culture : _context.Culture;

            if (chart.Kind == ChartKind.Pie)
            {
                return InsetAndClamp(area, area.Left + area.Width * 0.08f, area.Top + area.Height * 0.06f, area.Right - area.Width * 0.08f, area.Bottom - area.Height * 0.08f);
            }

            if (chart.Kind == ChartKind.Radar)
            {
                return InsetAndClamp(area, area.Left + area.Width * 0.05f, area.Top + area.Height * 0.08f, area.Right - area.Width * 0.09f, area.Bottom - area.Height * 0.12f);
            }

            if (chart.Series.Count == 0)
            {
                return InsetAndClamp(area, area.Left + area.Width * 0.18f, area.Top + area.Height * 0.18f, area.Right - area.Width * 0.18f, area.Bottom - area.Height * 0.18f);
            }

            var plot = ResolvePlotRect(area, chart);
            var left = plot.Left;
            var top = plot.Top;
            var right = plot.Right;
            var bottom = plot.Bottom;

            var axisLabelMaxWidth = MeasureMaxValueAxisLabelWidth(chart);
            left = Math.Min(left, plot.Left - 7f - axisLabelMaxWidth);

            if (!string.IsNullOrEmpty(chart.Title))
            {
                top = Math.Min(top, area.Top + area.Height * 0.06f);
            }

            if (chart.Categories.Count > 0)
            {
                bottom = Math.Max(bottom, plot.Bottom + LabelSizePt + 10f);
            }

            var dateAxisLayout = ResolveDateAxisLayout(chart, plot.Width);
            if (chart.HasDateCategoryAxis && Math.Abs(dateAxisLayout.RotationDeg) > 0.1f)
            {
                bottom = Math.Max(bottom, plot.Bottom + 44f);
            }
            else if (!chart.HasDateCategoryAxis && Math.Abs(ResolveStandardCategoryAxisRotation(chart, plot.Width)) > 0.1f)
            {
                bottom = Math.Max(bottom, plot.Bottom + 44f);
            }

            if (chart.LegendPosition != null && chart.Series.Count > 0)
            {
                var legendBounds = GetLegendBounds(area, plot, chart);
                left = Math.Min(left, legendBounds.Left);
                right = Math.Max(right, legendBounds.Right);
                bottom = Math.Max(bottom, legendBounds.Bottom);
            }

            return InsetAndClamp(area, left - 2f, top - 2f, right + 2f, bottom + 2f);
        }

        private SKRect ResolvePlotRect(SKRect area, ParsedChart chart)
        {
            if (chart.Kind == ChartKind.Radar)
            {
                var radarLeft = area.Left + area.Width * 0.11f;
                var radarTop = area.Top + (chart.Title != null ? area.Height * 0.1f : area.Height * 0.1f);
                var radarRight = area.Right - area.Width * 0.15f;
                var radarBottom = area.Bottom - area.Height * (chart.LegendPosition != null ? 0.10f : 0.12f);
                return new SKRect(radarLeft, radarTop, radarRight, radarBottom);
            }

            if (chart.HasManualPlot)
            {
                var l = area.Left + (float)(chart.PlotX * area.Width);
                var t = area.Top + (float)(chart.PlotY * area.Height);
                return new SKRect(l, t, l + (float)(chart.PlotW * area.Width), t + (float)(chart.PlotH * area.Height));
            }

            var axisLabelMaxWidth = MeasureMaxValueAxisLabelWidth(chart);
            var leftPadding = Math.Max(area.Width * 0.06f, axisLabelMaxWidth + 14f);
            var left = area.Left + leftPadding;
            var top = area.Top + (chart.Title != null ? area.Height * 0.205f : area.Height * 0.085f);
            var right = area.Right - area.Width * 0.07f;
            var bottomReserve = BottomReserveFraction(chart, area.Width);

            var dateAxisLayout = ResolveDateAxisLayout(chart, area.Width * 0.83f);
            if (chart.HasDateCategoryAxis
                && Math.Abs(dateAxisLayout.RotationDeg) > 0.1f
                && chart.LegendPosition != null)
            {
                // Excel keeps percent-stacked/date-axis line charts with rotated month labels a bit
                // tighter than our generic chart heuristic: the plot sits higher, is slightly
                // narrower, and leaves more room for the legend block below.
                left = area.Left + Math.Max(area.Width * 0.068f, axisLabelMaxWidth + 18f);
                top = area.Top + (chart.Title != null ? area.Height * 0.165f : area.Height * 0.095f);
                right = area.Right - area.Width * 0.078f;
                bottomReserve = Math.Max(bottomReserve, 0.315f);

                if (chart.Kind == ChartKind.Line && chart.IsStacked && !chart.IsPercentStacked)
                {
                    left = area.Left + Math.Max(area.Width * 0.064f, axisLabelMaxWidth + 17f);
                    top = area.Top + (chart.Title != null ? area.Height * 0.172f : area.Height * 0.1f);
                    right = area.Right - area.Width * 0.064f;
                    bottomReserve = Math.Max(bottomReserve, 0.323f);
                }
            }

            var bottom = area.Bottom - area.Height * bottomReserve;
            return new SKRect(left, top, right, bottom);
        }

        private void DrawTitle(SKCanvas canvas, PageLayout page, SKRect area, ParsedChart chart, float centerY)
        {
            if (string.IsNullOrEmpty(chart.Title))
            {
                return;
            }

            var clip = ResolveTitleBounds(area, chart, centerY);
            DrawTitleBorder(canvas, clip, chart);
            var titleFont = ResolveTitleFont(chart);
            DrawTextWithoutPdfOptimization(canvas, clip, chart.Title, area.MidX, centerY, titleFont, chart.TitleColor, TextAlign.Center, VAlign.Middle);
        }

        private SKRect ResolveTitleBounds(SKRect area, ParsedChart chart, float centerY)
        {
            var titleFont = ResolveTitleFont(chart);
            var textWidth = MeasureText(chart.Title, titleFont);
            var horizontalPadding = Math.Max(3.5f, area.Width * 0.007f);
            var verticalPadding = 0.25f;
            var width = Math.Min(area.Width * 0.34f, textWidth + horizontalPadding * 2f);
            var left = area.MidX - width * 0.5f;
            var right = area.MidX + width * 0.5f;

            using (var paint = TextPaint(titleFont, chart.TitleColor))
            {
                var metrics = paint.FontMetrics;
                var baseline = centerY - (metrics.Ascent + metrics.Descent) / 2f;
                var top = baseline + metrics.Ascent - verticalPadding;
                var bottom = baseline + metrics.Descent + verticalPadding;
                return new SKRect(left, top, right, bottom);
            }
        }

        private void DrawTitleBorder(SKCanvas canvas, SKRect bounds, ParsedChart chart)
        {
            if (canvas == null || chart == null || !chart.HasTitleBorder)
            {
                return;
            }

            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.TitleBorderColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.TitleBorderWidthPt),
                IsAntialias = false
            })
            {
                canvas.DrawRect(bounds, paint);
            }
        }

        private void DrawValueAxis(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart, double min, double max, double major)
        {
            if (max <= min || major <= 0d)
            {
                return;
            }

            major = ResolveRenderedMajorUnit(chart, plot, major);

            if (chart.HasMinorValueGridlines)
            {
                var minor = ResolveRenderedMinorUnit(chart, major);
                var gridlineRight = ResolveValueGridlineRight(plot, chart);
                if (minor > 0d && minor < major)
                {
                    using (var minorGrid = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        Color = chart.MinorValueGridlineColor,
                        StrokeWidth = (float)Math.Max(0.5d, chart.MinorValueGridlineWidthPt),
                        IsAntialias = false
                    })
                    {
                        var epsilon = minor * 0.1d;
                        for (var value = min + minor; value < max - epsilon; value += minor)
                        {
                            var ratio = value / major;
                            var nearestMajor = Math.Round(ratio);
                            if (Math.Abs(ratio - nearestMajor) <= 0.001d)
                            {
                                continue;
                            }

                            var y = ValueToY(value, min, max, plot);
                            canvas.DrawLine(plot.Left, y, gridlineRight, y, minorGrid);
                        }
                    }
                }
            }

            using (var grid = new SKPaint { Style = SKPaintStyle.Stroke, Color = GridColor, StrokeWidth = 1f, IsAntialias = false })
            {
                var gridlineRight = ResolveValueGridlineRight(plot, chart);
                var steps = (int)Math.Round((max - min) / major);
                for (var i = 0; i <= steps; i++)
                {
                    var value = min + i * major;
                    var y = ValueToY(value, min, max, plot);

                    if (chart.HasValueGridlines)
                    {
                        canvas.DrawLine(plot.Left, y, gridlineRight, y, grid);
                    }

                    var label = FormatValue(value, chart.ValueFormatCode);
                    var width = MeasureText(label, LabelSizePt);
                    var clip = new SKRect(plot.Left - 8f - width, y - LabelSizePt, plot.Left - 1f, y + LabelSizePt);
                    DrawText(canvas, page, clip, label, plot.Left - 5f, y, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Right, VAlign.Middle);
                }
            }
        }

        private void DrawCategoryAxis(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            using (var axis = new SKPaint { Style = SKPaintStyle.Stroke, Color = GridColor, StrokeWidth = 1f, IsAntialias = false })
            {
                canvas.DrawLine(plot.Left, plot.Bottom, plot.Right, plot.Bottom, axis);
            }

            if (chart.HasDateCategoryAxis)
            {
                DrawDateCategoryAxis(canvas, page, area, plot, chart);
                return;
            }

            if (_suppressCategoryAxisLabels)
            {
                return;
            }

            var count = chart.Categories.Count;
            var visibleSlice = ResolveVisibleChartSlice(page, area);
            var rotatedClip = new SKRect(
                Math.Max(plot.Left - 36f, visibleSlice.Left),
                plot.Bottom + 2f,
                Math.Min(plot.Right + 18f, visibleSlice.Right),
                visibleSlice.Bottom);
            var rotationDeg = ResolveStandardCategoryAxisRotation(chart, plot.Width);
            for (var i = 0; i < count; i++)
            {
                var x = CategoryX(i, count, plot, chart.CrossBetween);
                var label = chart.Categories[i];
                if (Math.Abs(rotationDeg) > 0.1f)
                {
                    if (x < visibleSlice.Left - 10f || x > visibleSlice.Right + 24f)
                    {
                        continue;
                    }

                    DrawRotatedText(canvas, page, rotatedClip, label, x + 2f, plot.Bottom + 6f, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Right, VAlign.Top, rotationDeg);
                }
                else
                {
                    var width = MeasureText(label, LabelSizePt);
                    var clip = new SKRect(x - width * 0.6f, plot.Bottom + 2f, x + width * 0.6f, plot.Bottom + LabelSizePt * 2.2f);
                    DrawText(canvas, page, clip, label, x, plot.Bottom + 4f, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Center, VAlign.Top);
                }
            }
        }

        private float ResolveStandardCategoryAxisRotation(ParsedChart chart, float availableWidth)
        {
            if (chart == null || chart.HasDateCategoryAxis || Math.Abs(chart.CategoryAxisTextRotationDeg) <= 0.1d)
            {
                return 0f;
            }

            var count = chart.Categories != null ? chart.Categories.Count : 0;
            if (count <= 0)
            {
                return 0f;
            }

            var slotWidth = availableWidth / count;
            var widestLabel = 0f;
            for (var i = 0; i < chart.Categories.Count; i++)
            {
                var label = chart.Categories[i] ?? string.Empty;
                widestLabel = Math.Max(widestLabel, MeasureText(label, LabelSizePt));
            }

            if (widestLabel <= slotWidth * 0.9f)
            {
                return 0f;
            }

            return (float)chart.CategoryAxisTextRotationDeg;
        }

        private void DrawSeries(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max)
        {
            var pointCount = MaxPointCount(chart);

            if (chart.Kind == ChartKind.Radar)
            {
                return;
            }

            if (chart.Kind == ChartKind.Column)
            {
                DrawColumns(canvas, plot, chart, min, max, pointCount);
                return;
            }

            if (chart.Kind == ChartKind.Bar)
            {
                DrawBars(canvas, plot, chart, min, max, pointCount);
                return;
            }

            var baselineY = ValueToY(Math.Max(min, 0d), min, max, plot);
            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
            {
                var series = chart.Series[seriesIndex];
                var displayValues = ChartStackingMath.BuildDisplayValues(chart, seriesIndex);
                var renderPoints = new List<SKPoint>();
                using (var line = new SKPaint { Style = SKPaintStyle.Stroke, Color = series.LineColor, StrokeWidth = (float)series.LineWidthPt, IsAntialias = true, StrokeCap = SKStrokeCap.Round, StrokeJoin = SKStrokeJoin.Round })
                {
                    var path = new SKPath();
                    var started = false;
                    var datePoints = ChartSeriesPointBuilder.BuildDateSeriesPoints(chart, displayValues, plot, min, max, _dateSystem);
                    if (datePoints != null)
                    {
                        for (var i = 0; i < datePoints.Count; i++)
                        {
                            var point = datePoints[i];
                            renderPoints.Add(point);
                            if (!started)
                            {
                                path.MoveTo(point.X, point.Y);
                                started = true;
                            }
                            else
                            {
                                path.LineTo(point.X, point.Y);
                            }
                        }
                    }
                    else
                    {
                        for (var i = 0; i < displayValues.Count; i++)
                        {
                            if (!displayValues[i].HasValue)
                            {
                                started = false;
                                continue;
                            }

                            var x = PointX(chart, i, pointCount, plot);
                            var y = ValueToY(displayValues[i].Value, min, max, plot);
                            renderPoints.Add(new SKPoint(x, y));
                            if (!started) { path.MoveTo(x, y); started = true; }
                            else path.LineTo(x, y);
                        }
                    }

                    if (chart.Kind == ChartKind.Area)
                    {
                        using (var fill = new SKPaint
                        {
                            Style = SKPaintStyle.Fill,
                            Color = new SKColor(series.Color.Red, series.Color.Green, series.Color.Blue, 255),
                            IsAntialias = true
                        })
                        {
                            using (var areaPath = BuildAreaFillPath(chart, plot, min, max, pointCount, seriesIndex, displayValues, renderPoints, baselineY))
                            {
                                if (areaPath != null)
                                {
                                    canvas.DrawPath(areaPath, fill);
                                }
                            }
                        }
                    }

                    if (series.HasVisibleLine)
                    {
                        canvas.DrawPath(path, line);
                    }
                    path.Dispose();
                }

                DrawSeriesMarkers(canvas, series, renderPoints);
            }
        }

        private SKPath BuildAreaFillPath(ParsedChart chart, SKRect plot, double min, double max, int pointCount, int seriesIndex, List<double?> upperValues, List<SKPoint> upperPoints, float baselineY)
        {
            if (upperPoints == null || upperPoints.Count == 0)
            {
                return null;
            }

            var areaPath = new SKPath();
            for (var i = 0; i < upperPoints.Count; i++)
            {
                var point = upperPoints[i];
                if (i == 0)
                {
                    areaPath.MoveTo(point.X, point.Y);
                }
                else
                {
                    areaPath.LineTo(point.X, point.Y);
                }
            }

            List<SKPoint> lowerPoints;
            if (chart != null && (chart.IsStacked || chart.IsPercentStacked) && seriesIndex > 0)
            {
                var lowerValues = ChartStackingMath.BuildDisplayValues(chart, seriesIndex - 1);
                lowerPoints = ChartSeriesPointBuilder.BuildDateSeriesPoints(chart, lowerValues, plot, min, max, _dateSystem);
                if (lowerPoints == null)
                {
                    lowerPoints = BuildCategorySeriesPoints(chart, lowerValues, pointCount, plot, min, max);
                }
            }
            else
            {
                lowerPoints = BuildAreaBaselinePoints(chart, upperValues, upperPoints, pointCount, plot, baselineY);
            }

            if (lowerPoints == null || lowerPoints.Count == 0)
            {
                areaPath.Dispose();
                return null;
            }

            for (var i = lowerPoints.Count - 1; i >= 0; i--)
            {
                areaPath.LineTo(lowerPoints[i].X, lowerPoints[i].Y);
            }

            areaPath.Close();
            return areaPath;
        }

        private List<SKPoint> BuildCategorySeriesPoints(ParsedChart chart, List<double?> values, int pointCount, SKRect plot, double min, double max)
        {
            var points = new List<SKPoint>();
            if (values == null)
            {
                return points;
            }

            for (var i = 0; i < values.Count; i++)
            {
                if (!values[i].HasValue)
                {
                    continue;
                }

                var x = PointX(chart, i, pointCount, plot);
                var y = ValueToY(values[i].Value, min, max, plot);
                points.Add(new SKPoint(x, y));
            }

            return points;
        }

        private void DrawRadar(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart, double min, double max, double major)
        {
            if (canvas == null || chart == null || chart.Series == null || chart.Series.Count == 0 || max <= min)
            {
                return;
            }

            var categoryCount = chart.Categories != null ? chart.Categories.Count : 0;
            if (categoryCount <= 2)
            {
                return;
            }

            major = ResolveRenderedMajorUnit(chart, plot, major);
            var center = new SKPoint(plot.MidX, plot.MidY + 1f);
            var radius = ResolveRadarRadius(plot);

            DrawRadarGrid(canvas, chart, center, radius, min, max, major);
            DrawRadarCategoryLabels(canvas, page, area, plot, chart, center, radius);
            DrawRadarValueAxisLabels(canvas, page, area, plot, chart, center, radius, min, max, major);

            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
            {
                var series = chart.Series[seriesIndex];
                var points = BuildRadarSeriesPoints(chart, series, center, radius, min, max);
                if (points.Count == 0)
                {
                    continue;
                }

                using (var path = BuildRadarPath(points))
                {
                    if (path == null)
                    {
                        continue;
                    }

                    if (string.Equals(chart.RadarStyle, "filled", StringComparison.Ordinal))
                    {
                        using (var fill = new SKPaint
                        {
                            Style = SKPaintStyle.Fill,
                            Color = new SKColor(series.Color.Red, series.Color.Green, series.Color.Blue, 110),
                            IsAntialias = true
                        })
                        {
                            canvas.DrawPath(path, fill);
                        }
                    }

                    if (series.HasVisibleLine)
                    {
                        using (var line = new SKPaint
                        {
                            Style = SKPaintStyle.Stroke,
                            Color = series.LineColor,
                            StrokeWidth = (float)series.LineWidthPt,
                            IsAntialias = true,
                            StrokeCap = SKStrokeCap.Round,
                            StrokeJoin = SKStrokeJoin.Round
                        })
                        {
                            canvas.DrawPath(path, line);
                        }
                    }
                }

                DrawSeriesMarkers(canvas, series, points);
            }
        }

        private float ResolveRadarRadius(SKRect plot)
        {
            var radius = Math.Min(plot.Width, plot.Height) * 0.462f;
            if (radius < 12f)
            {
                radius = 12f;
            }

            return radius;
        }

        private void DrawRadarGrid(SKCanvas canvas, ParsedChart chart, SKPoint center, float radius, double min, double max, double major)
        {
            var categoryCount = chart.Categories != null ? chart.Categories.Count : 0;
            if (categoryCount <= 2)
            {
                return;
            }

            using (var spokePaint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.HasCategoryGridlines ? chart.CategoryGridlineColor : GridColor,
                StrokeWidth = 1f,
                IsAntialias = false
            })
            using (var ringPaint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.HasValueGridlines ? GridColor : chart.CategoryGridlineColor,
                StrokeWidth = 1f,
                IsAntialias = false
            })
            {
                for (var index = 0; index < categoryCount; index++)
                {
                    var angle = RadarAngle(index, categoryCount);
                    var outerPoint = RadarPoint(center, radius, angle);
                    canvas.DrawLine(center.X, center.Y, outerPoint.X, outerPoint.Y, spokePaint);
                }

                if (major <= 0d)
                {
                    major = (max - min) / 5d;
                }

                if (major <= 0d)
                {
                    return;
                }

                var steps = (int)Math.Round((max - min) / major);
                if (steps <= 0)
                {
                    steps = 1;
                }

                for (var step = 1; step <= steps; step++)
                {
                    var value = min + step * major;
                    if (value > max + major * 0.001d)
                    {
                        value = max;
                    }

                    var currentRadius = RadarRadiusForValue(value, min, max, radius);
                    using (var path = new SKPath())
                    {
                        for (var index = 0; index < categoryCount; index++)
                        {
                            var angle = RadarAngle(index, categoryCount);
                            var point = RadarPoint(center, currentRadius, angle);
                            if (index == 0)
                            {
                                path.MoveTo(point.X, point.Y);
                            }
                            else
                            {
                                path.LineTo(point.X, point.Y);
                            }
                        }

                        path.Close();
                        canvas.DrawPath(path, ringPaint);
                    }
                }
            }
        }

        private void DrawRadarCategoryLabels(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart, SKPoint center, float radius)
        {
            if (_suppressCategoryAxisLabels || chart.Categories == null)
            {
                return;
            }

            var count = chart.Categories.Count;
            var labelRadius = radius + 5f;
            for (var index = 0; index < count; index++)
            {
                var angle = RadarAngle(index, count);
                var point = RadarPoint(center, labelRadius, angle);
                var label = chart.Categories[index] ?? string.Empty;
                var width = MeasureText(label, LabelSizePt);
                var clip = new SKRect(point.X - width, point.Y - LabelSizePt * 1.1f, point.X + width, point.Y + LabelSizePt * 1.1f);
                DrawText(canvas, page, clip, label, point.X, point.Y, LabelSizePt, false, chart.AxisLabelColor, RadarTextAlign(angle), RadarVerticalAlign(angle));
            }
        }

        private void DrawRadarValueAxisLabels(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart, SKPoint center, float radius, double min, double max, double major)
        {
            if (max <= min)
            {
                return;
            }

            if (major <= 0d)
            {
                major = (max - min) / 5d;
            }

            if (major <= 0d)
            {
                return;
            }

            var steps = (int)Math.Round((max - min) / major);
            if (steps <= 0)
            {
                steps = 1;
            }

            for (var step = 0; step <= steps; step++)
            {
                var value = min + step * major;
                if (step == steps || value > max)
                {
                    value = max;
                }

                var currentRadius = RadarRadiusForValue(value, min, max, radius);
                var label = FormatValue(value, chart.ValueFormatCode);
                var width = MeasureText(label, LabelSizePt);
                var x = center.X - 8f;
                var y = center.Y - currentRadius;
                var clip = new SKRect(x - width - 4f, y - LabelSizePt, x + 1f, y + LabelSizePt);
                DrawText(canvas, page, clip, label, x, y, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Right, VAlign.Middle);
            }
        }

        private List<SKPoint> BuildRadarSeriesPoints(ParsedChart chart, ChartSeries series, SKPoint center, float radius, double min, double max)
        {
            var points = new List<SKPoint>();
            if (chart == null || series == null || chart.Categories == null)
            {
                return points;
            }

            var count = chart.Categories.Count;
            for (var index = 0; index < count; index++)
            {
                if (index >= series.Values.Count || !series.Values[index].HasValue)
                {
                    continue;
                }

                var angle = RadarAngle(index, count);
                var currentRadius = RadarRadiusForValue(series.Values[index].Value, min, max, radius);
                points.Add(RadarPoint(center, currentRadius, angle));
            }

            return points;
        }

        private SKPath BuildRadarPath(List<SKPoint> points)
        {
            if (points == null || points.Count == 0)
            {
                return null;
            }

            var path = new SKPath();
            for (var index = 0; index < points.Count; index++)
            {
                if (index == 0)
                {
                    path.MoveTo(points[index].X, points[index].Y);
                }
                else
                {
                    path.LineTo(points[index].X, points[index].Y);
                }
            }

            path.Close();
            return path;
        }

        private static float RadarAngle(int index, int count)
        {
            if (count <= 0)
            {
                return 0f;
            }

            return (float)(-Math.PI / 2d + Math.PI * 2d * index / count);
        }

        private static SKPoint RadarPoint(SKPoint center, float radius, float angle)
        {
            return new SKPoint(center.X + radius * (float)Math.Cos(angle), center.Y + radius * (float)Math.Sin(angle));
        }

        private static float RadarRadiusForValue(double value, double min, double max, float outerRadius)
        {
            if (max <= min)
            {
                return 0f;
            }

            var ratio = (value - min) / (max - min);
            if (ratio < 0d)
            {
                ratio = 0d;
            }
            else if (ratio > 1d)
            {
                ratio = 1d;
            }

            return (float)(ratio * outerRadius);
        }

        private TextAlign RadarTextAlign(float angle)
        {
            var cosine = Math.Cos(angle);
            if (cosine > 0.25d)
            {
                return TextAlign.Left;
            }

            if (cosine < -0.25d)
            {
                return TextAlign.Right;
            }

            return TextAlign.Center;
        }

        private VAlign RadarVerticalAlign(float angle)
        {
            var sine = Math.Sin(angle);
            if (sine > 0.35d)
            {
                return VAlign.Top;
            }

            if (sine < -0.35d)
            {
                return VAlign.Baseline;
            }

            return VAlign.Middle;
        }

        private List<SKPoint> BuildAreaBaselinePoints(ParsedChart chart, List<double?> values, List<SKPoint> upperPoints, int pointCount, SKRect plot, float baselineY)
        {
            var points = new List<SKPoint>();
            if (upperPoints == null || upperPoints.Count == 0)
            {
                return points;
            }

            if (chart != null && chart.HasDateCategoryAxis && chart.CategoryValues != null)
            {
                var count = chart.CategoryValues.Count;
                if (values != null && values.Count < count)
                {
                    count = values.Count;
                }

                for (var i = 0; i < count; i++)
                {
                    if (!chart.CategoryValues[i].HasValue || values == null || !values[i].HasValue)
                    {
                        continue;
                    }

                    var serial = ChartDateAxisMath.NormalizePointSerial(chart, chart.CategoryValues[i].Value, _dateSystem);
                    double plotMinDate;
                    double plotMaxDate;
                    if (!ChartDateAxisMath.TryGetDateAxisPlotRange(chart, _dateSystem, out plotMinDate, out plotMaxDate))
                    {
                        break;
                    }

                    points.Add(new SKPoint(
                        ChartDateAxisMath.DateCategoryX(serial, plotMinDate, plotMaxDate, plot),
                        baselineY));
                }

                if (points.Count > 0)
                {
                    return points;
                }
            }

            for (var i = 0; i < upperPoints.Count; i++)
            {
                points.Add(new SKPoint(upperPoints[i].X, baselineY));
            }

            return points;
        }

        private void DrawSeriesMarkers(SKCanvas canvas, ChartSeries series, List<SKPoint> points)
        {
            if (canvas == null || series == null || points == null || points.Count == 0 || !series.MarkerVisible)
            {
                return;
            }

            var diameter = Math.Max(3f, series.MarkerSize);
            var radius = diameter * 0.5f;
            var strokeWidth = (float)Math.Max(0.75d, series.MarkerStrokeWidthPt);

            using (var fill = new SKPaint
            {
                Style = SKPaintStyle.Fill,
                Color = series.MarkerFillColor,
                IsAntialias = true
            })
            using (var stroke = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = series.MarkerStrokeColor,
                StrokeWidth = strokeWidth,
                IsAntialias = true
            })
            {
                for (var i = 0; i < points.Count; i++)
                {
                    DrawMarker(canvas, points[i], radius, series.MarkerSymbol, fill, stroke);
                }
            }
        }

        private void DrawMarker(SKCanvas canvas, SKPoint center, float radius, string symbol, SKPaint fill, SKPaint stroke)
        {
            if (string.Equals(symbol, "none", StringComparison.Ordinal))
            {
                return;
            }

            if (string.Equals(symbol, "square", StringComparison.Ordinal))
            {
                var rect = new SKRect(center.X - radius, center.Y - radius, center.X + radius, center.Y + radius);
                canvas.DrawRect(rect, fill);
                canvas.DrawRect(rect, stroke);
                return;
            }

            if (string.Equals(symbol, "diamond", StringComparison.Ordinal))
            {
                using (var path = new SKPath())
                {
                    path.MoveTo(center.X, center.Y - radius);
                    path.LineTo(center.X + radius, center.Y);
                    path.LineTo(center.X, center.Y + radius);
                    path.LineTo(center.X - radius, center.Y);
                    path.Close();
                    canvas.DrawPath(path, fill);
                    canvas.DrawPath(path, stroke);
                }
                return;
            }

            if (string.Equals(symbol, "triangle", StringComparison.Ordinal))
            {
                using (var path = new SKPath())
                {
                    path.MoveTo(center.X, center.Y - radius);
                    path.LineTo(center.X + radius * 0.9f, center.Y + radius * 0.8f);
                    path.LineTo(center.X - radius * 0.9f, center.Y + radius * 0.8f);
                    path.Close();
                    canvas.DrawPath(path, fill);
                    canvas.DrawPath(path, stroke);
                }
                return;
            }

            canvas.DrawCircle(center, radius, fill);
            canvas.DrawCircle(center, radius, stroke);
        }

        private void DrawColumns(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, int pointCount)
        {
            if (pointCount == 0 || chart.Series.Count == 0)
            {
                return;
            }

            var slot = plot.Width / pointCount;
            var columnGeometry = ResolveColumnGeometry(chart);
            var groupWidth = slot * columnGeometry.Item1;
            var barWidth = slot * columnGeometry.Item2;
            var seriesPitch = slot * columnGeometry.Item3;
            var baselineY = ValueToY(Math.Max(min, 0d), min, max, plot);

            if (chart.IsStacked || chart.IsPercentStacked)
            {
                DrawStackedColumns(canvas, plot, chart, min, max, pointCount, slot, groupWidth, baselineY);
                return;
            }

            for (var i = 0; i < pointCount; i++)
            {
                var slotLeft = plot.Left + i * slot + (slot - groupWidth) / 2f;
                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    if (i >= series.Values.Count || !series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var y = ValueToY(series.Values[i].Value, min, max, plot);
                    var left = slotLeft + s * seriesPitch;
                    var top = Math.Min(y, baselineY);
                    var bottom = Math.Max(y, baselineY);
                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = series.Color, IsAntialias = true })
                    {
                        canvas.DrawRect(new SKRect(left, top, left + barWidth, bottom), fill);
                    }
                }
            }
        }

        private void DrawBars(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, int pointCount)
        {
            if (pointCount == 0 || chart.Series.Count == 0)
            {
                return;
            }

            var slot = plot.Height / pointCount;
            var barGeometry = ResolveColumnGeometry(chart);
            var groupHeight = slot * barGeometry.Item1;
            var barHeight = slot * barGeometry.Item2;
            var seriesPitch = slot * barGeometry.Item3;
            var baselineX = ValueToX(Math.Max(min, 0d), min, max, plot);

            if (chart.IsStacked || chart.IsPercentStacked)
            {
                DrawStackedBars(canvas, plot, chart, min, max, pointCount, slot, groupHeight, baselineX);
                return;
            }

            for (var i = 0; i < pointCount; i++)
            {
                var categoryIndex = ResolveBarCategoryIndex(chart, i, pointCount);
                var slotTop = plot.Top + categoryIndex * slot + (slot - groupHeight) / 2f;
                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    if (i >= series.Values.Count || !series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var x = ValueToX(series.Values[i].Value, min, max, plot);
                    var top = slotTop + (chart.Series.Count - 1 - s) * seriesPitch;
                    var left = Math.Min(x, baselineX);
                    var right = Math.Max(x, baselineX);
                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = series.Color, IsAntialias = true })
                    {
                        canvas.DrawRect(new SKRect(left, top, right, top + barHeight), fill);
                    }
                }
            }
        }

        private void DrawStackedBars(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, int pointCount, float slot, float barHeight, float baselineX)
        {
            var positiveTotals = new double[pointCount];
            var negativeTotals = new double[pointCount];

            if (chart.IsPercentStacked)
            {
                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    var count = Math.Min(pointCount, series.Values.Count);
                    for (var i = 0; i < count; i++)
                    {
                        if (!series.Values[i].HasValue)
                        {
                            continue;
                        }

                        var value = series.Values[i].Value;
                        if (value >= 0d)
                        {
                            positiveTotals[i] += value;
                        }
                        else
                        {
                            negativeTotals[i] += Math.Abs(value);
                        }
                    }
                }
            }

            var positiveStarts = new double[pointCount];
            var negativeStarts = new double[pointCount];
            for (var i = 0; i < pointCount; i++)
            {
                var categoryIndex = ResolveBarCategoryIndex(chart, i, pointCount);
                var top = plot.Top + categoryIndex * slot + (slot - barHeight) / 2f;
                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    if (i >= series.Values.Count || !series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var value = series.Values[i].Value;
                    if (chart.IsPercentStacked)
                    {
                        if (value >= 0d)
                        {
                            if (positiveTotals[i] <= 0d)
                            {
                                continue;
                            }

                            value /= positiveTotals[i];
                        }
                        else
                        {
                            if (negativeTotals[i] <= 0d)
                            {
                                continue;
                            }

                            value /= negativeTotals[i];
                        }
                    }

                    double startValue;
                    double endValue;
                    if (value >= 0d)
                    {
                        startValue = positiveStarts[i];
                        endValue = startValue + value;
                        positiveStarts[i] = endValue;
                    }
                    else
                    {
                        startValue = negativeStarts[i];
                        endValue = startValue + value;
                        negativeStarts[i] = endValue;
                    }

                    var startX = ValueToX(startValue, min, max, plot);
                    var endX = ValueToX(endValue, min, max, plot);
                    var left = Math.Min(startX, endX);
                    var right = Math.Max(startX, endX);
                    if (Math.Abs(right - left) < 0.1f)
                    {
                        left = Math.Min(left, baselineX);
                        right = Math.Max(right, baselineX);
                    }

                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = series.Color, IsAntialias = true })
                    {
                        canvas.DrawRect(new SKRect(left, top, right, top + barHeight), fill);
                    }
                }
            }
        }

        private void DrawStackedColumns(SKCanvas canvas, SKRect plot, ParsedChart chart, double min, double max, int pointCount, float slot, float barWidth, float baselineY)
        {
            var positiveTotals = new double[pointCount];
            var negativeTotals = new double[pointCount];

            if (chart.IsPercentStacked)
            {
                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    var count = Math.Min(pointCount, series.Values.Count);
                    for (var i = 0; i < count; i++)
                    {
                        if (!series.Values[i].HasValue)
                        {
                            continue;
                        }

                        var value = series.Values[i].Value;
                        if (value >= 0d)
                        {
                            positiveTotals[i] += value;
                        }
                        else
                        {
                            negativeTotals[i] += Math.Abs(value);
                        }
                    }
                }
            }

            var positiveStarts = new double[pointCount];
            var negativeStarts = new double[pointCount];
            for (var i = 0; i < pointCount; i++)
            {
                var left = plot.Left + i * slot + (slot - barWidth) / 2f;
                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    if (i >= series.Values.Count || !series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var value = series.Values[i].Value;
                    if (chart.IsPercentStacked)
                    {
                        if (value >= 0d)
                        {
                            if (positiveTotals[i] <= 0d)
                            {
                                continue;
                            }

                            value /= positiveTotals[i];
                        }
                        else
                        {
                            if (negativeTotals[i] <= 0d)
                            {
                                continue;
                            }

                            value /= negativeTotals[i];
                        }
                    }

                    double startValue;
                    double endValue;
                    if (value >= 0d)
                    {
                        startValue = positiveStarts[i];
                        endValue = startValue + value;
                        positiveStarts[i] = endValue;
                    }
                    else
                    {
                        startValue = negativeStarts[i];
                        endValue = startValue + value;
                        negativeStarts[i] = endValue;
                    }

                    var startY = ValueToY(startValue, min, max, plot);
                    var endY = ValueToY(endValue, min, max, plot);
                    var top = Math.Min(startY, endY);
                    var bottom = Math.Max(startY, endY);
                    if (Math.Abs(bottom - top) < 0.1f)
                    {
                        top = Math.Min(top, baselineY);
                        bottom = Math.Max(bottom, baselineY);
                    }

                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = series.Color, IsAntialias = true })
                    {
                        canvas.DrawRect(new SKRect(left, top, left + barWidth, bottom), fill);
                    }
                }
            }
        }

        private static Tuple<float, float, float> ResolveColumnGeometry(ParsedChart chart)
        {
            var seriesCount = chart != null && chart.Series != null ? chart.Series.Count : 0;
            if (seriesCount <= 0)
            {
                return Tuple.Create(0f, 0f, 0f);
            }

            if (chart == null)
            {
                var legacyGroupWidth = 0.8f;
                var legacyBarWidth = legacyGroupWidth / seriesCount;
                return Tuple.Create(legacyGroupWidth, legacyBarWidth, legacyBarWidth);
            }

            var gapRatio = (float)(chart.GapWidthPercent / 100d);
            if (gapRatio < 0f)
            {
                gapRatio = 0f;
            }

            if (chart.IsStacked || chart.IsPercentStacked)
            {
                var stackedWidthRatio = 1f / (1f + gapRatio);
                return Tuple.Create(stackedWidthRatio, stackedWidthRatio, 0f);
            }

            var overlapRatio = (float)(chart.OverlapPercent / 100d);
            if (overlapRatio < -1f)
            {
                overlapRatio = -1f;
            }
            else if (overlapRatio > 1f)
            {
                overlapRatio = 1f;
            }

            var pitchFactor = 1f - overlapRatio;
            if (pitchFactor < 0f)
            {
                pitchFactor = 0f;
            }

            var spanFactor = 1f + (seriesCount - 1) * pitchFactor;
            var denominator = spanFactor + gapRatio;
            if (denominator <= 0f)
            {
                var fallbackGroupWidth = 0.8f;
                var fallbackBarWidth = fallbackGroupWidth / seriesCount;
                return Tuple.Create(fallbackGroupWidth, fallbackBarWidth, fallbackBarWidth);
            }

            var barWidthRatio = 1f / denominator;
            var seriesPitchRatio = barWidthRatio * pitchFactor;
            var groupWidthRatio = barWidthRatio * spanFactor;
            return Tuple.Create(groupWidthRatio, barWidthRatio, seriesPitchRatio);
        }

        private void DrawLegend(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            var legendItemCount = LegendItemCount(chart);
            if (chart.LegendPosition == null || legendItemCount == 0)
            {
                return;
            }

            // Only the common bottom legend is laid out; others fall back to the bottom band.
            var y = LegendCenterY(area, plot, chart);
            const float swatch = 9f;
            var gap = ResolveLegendTextGap(chart);
            var itemGap = ResolveLegendItemGap(chart);

            var widths = new float[legendItemCount];
            var total = 0f;
            for (var i = 0; i < legendItemCount; i++)
            {
                widths[i] = swatch + gap + MeasureText(LegendItemLabel(chart, i), LabelSizePt);
                total += widths[i] + (i > 0 ? itemGap : 0f);
            }

            var x = area.MidX - total / 2f;
            var legendTextBottom = y;
            var legendSwatchHeight = 5.5f;
            using (var metricsPaint = TextPaint(ChartFont(LabelSizePt, false), chart.LegendTextColor))
            {
                var metrics = metricsPaint.FontMetrics;
                legendTextBottom = y + (metrics.Descent - metrics.Ascent) / 2f - 1.4f;
            }

            for (var i = 0; i < legendItemCount; i++)
            {
                var label = LegendItemLabel(chart, i);
                var color = LegendItemColor(chart, i);
                if (chart.Kind == ChartKind.Column || chart.Kind == ChartKind.Bar || chart.Kind == ChartKind.Pie || chart.Kind == ChartKind.Area)
                {
                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = color, IsAntialias = true })
                    {
                        var rectBottom = legendTextBottom;
                        var rectTop = rectBottom - legendSwatchHeight;
                        canvas.DrawRect(new SKRect(x, rectTop, x + legendSwatchHeight, rectBottom), fill);
                    }
                }
                else if (chart.Kind == ChartKind.Radar && string.Equals(chart.RadarStyle, "filled", StringComparison.Ordinal))
                {
                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = new SKColor(color.Red, color.Green, color.Blue, 110), IsAntialias = true })
                    using (var stroke = new SKPaint { Style = SKPaintStyle.Stroke, Color = color, StrokeWidth = 1.4f, IsAntialias = true })
                    {
                        var rectBottom = legendTextBottom;
                        var rectTop = rectBottom - legendSwatchHeight;
                        var rect = new SKRect(x, rectTop, x + legendSwatchHeight, rectBottom);
                        canvas.DrawRect(rect, fill);
                        canvas.DrawRect(rect, stroke);
                    }
                }
                else
                {
                    using (var line = new SKPaint { Style = SKPaintStyle.Stroke, Color = color, StrokeWidth = 2.5f, IsAntialias = true, StrokeCap = SKStrokeCap.Round })
                    {
                        canvas.DrawLine(x, y, x + swatch, y, line);
                    }

                    if (chart.Kind == ChartKind.Radar && string.Equals(chart.RadarStyle, "marker", StringComparison.Ordinal))
                    {
                        var seriesIndex = ResolveLegendSeriesIndex(chart, i);
                        if (chart.Series != null && seriesIndex >= 0 && seriesIndex < chart.Series.Count)
                        {
                            using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = chart.Series[seriesIndex].MarkerFillColor, IsAntialias = true })
                            using (var stroke = new SKPaint { Style = SKPaintStyle.Stroke, Color = chart.Series[seriesIndex].MarkerStrokeColor, StrokeWidth = 1f, IsAntialias = true })
                            {
                                DrawMarker(canvas, new SKPoint(x + swatch * 0.5f, y), 2.35f, chart.Series[seriesIndex].MarkerSymbol, fill, stroke);
                            }
                        }
                    }
                }
                var labelWidth = MeasureText(label, LabelSizePt);
                var clip = new SKRect(x + swatch + gap, y - LabelSizePt, x + swatch + gap + labelWidth + 2f, y + LabelSizePt);
                DrawText(canvas, page, clip, label, x + swatch + gap, y, LabelSizePt, false, chart.LegendTextColor, TextAlign.Left, VAlign.Middle);
                x += widths[i] + itemGap;
            }
        }

        private SKRect GetLegendBounds(SKRect area, SKRect plot, ParsedChart chart)
        {
            var y = LegendCenterY(area, plot, chart);
            const float swatch = 9f;
            var gap = ResolveLegendTextGap(chart);
            var itemGap = ResolveLegendItemGap(chart);

            var total = 0f;
            var legendItemCount = LegendItemCount(chart);
            for (var i = 0; i < legendItemCount; i++)
            {
                total += swatch + gap + MeasureText(LegendItemLabel(chart, i), LabelSizePt);
                if (i > 0)
                {
                    total += itemGap;
                }
            }

            var x = area.MidX - total / 2f;
            return new SKRect(x, y - 8f, x + total, y + 8f);
        }

        private float BottomReserveFraction(ParsedChart chart, float availableWidth)
        {
            var bottomReserve = chart.LegendPosition != null ? 0.2f : 0.12f;
            var dateAxisLayout = ResolveDateAxisLayout(chart, availableWidth * 0.83f);
            if (chart.HasDateCategoryAxis && Math.Abs(dateAxisLayout.RotationDeg) > 0.1f)
            {
                if (chart.LegendPosition != null)
                {
                    bottomReserve = 0.285f;
                }
                else
                {
                    bottomReserve = 0.205f;
                }
            }

            return bottomReserve;
        }

        private int ResolveLegendSeriesIndex(ParsedChart chart, int legendIndex)
        {
            if (chart == null || chart.Series == null || legendIndex < 0 || legendIndex >= chart.Series.Count)
            {
                return legendIndex;
            }

            if (chart.Kind == ChartKind.Bar && !chart.IsStacked && !chart.IsPercentStacked)
            {
                return chart.Series.Count - 1 - legendIndex;
            }

            return legendIndex;
        }

        private int LegendItemCount(ParsedChart chart)
        {
            if (chart == null)
            {
                return 0;
            }

            if (chart.Kind == ChartKind.Pie)
            {
                return chart.Categories != null ? chart.Categories.Count : 0;
            }

            return chart.Series != null ? chart.Series.Count : 0;
        }

        private string LegendItemLabel(ParsedChart chart, int legendIndex)
        {
            if (chart == null || legendIndex < 0)
            {
                return string.Empty;
            }

            if (chart.Kind == ChartKind.Pie)
            {
                if (chart.Categories != null && legendIndex < chart.Categories.Count)
                {
                    return chart.Categories[legendIndex] ?? string.Empty;
                }

                return string.Empty;
            }

            var seriesIndex = ResolveLegendSeriesIndex(chart, legendIndex);
            if (chart.Series == null || seriesIndex < 0 || seriesIndex >= chart.Series.Count)
            {
                return string.Empty;
            }

            return chart.Series[seriesIndex].Name ?? string.Empty;
        }

        private SKColor LegendItemColor(ParsedChart chart, int legendIndex)
        {
            if (chart == null || legendIndex < 0)
            {
                return SKColors.Gray;
            }

            if (chart.Kind == ChartKind.Pie)
            {
                return _context.Colors.ResolveSchemeName("accent" + ((legendIndex % 6) + 1), SKColors.Gray);
            }

            var seriesIndex = ResolveLegendSeriesIndex(chart, legendIndex);
            if (chart.Series == null || seriesIndex < 0 || seriesIndex >= chart.Series.Count)
            {
                return SKColors.Gray;
            }

            return chart.Series[seriesIndex].Color;
        }

        private float LegendCenterY(SKRect area, SKRect plot, ParsedChart chart)
        {
            var position = 0.68f;
            var dateAxisLayout = ResolveDateAxisLayout(chart, plot.Width);
            if (chart.HasDateCategoryAxis && Math.Abs(dateAxisLayout.RotationDeg) > 0.1f)
            {
                position = 0.86f;
            }

            return plot.Bottom + (area.Bottom - plot.Bottom) * position;
        }

        private float ResolveLegendTextGap(ParsedChart chart)
        {
            if (chart != null && chart.Kind == ChartKind.Pie)
            {
                return 0f;
            }

            if (chart != null && chart.Kind == ChartKind.Area)
            {
                return 0f;
            }

            return 1f;
        }

        private float ResolveLegendItemGap(ParsedChart chart)
        {
            if (chart != null && chart.Kind == ChartKind.Pie)
            {
                return 7.2f;
            }

            if (chart != null && chart.Kind == ChartKind.Area)
            {
                return 6f;
            }

            return 12.6f;
        }

        private void DrawPie(SKCanvas canvas, PageLayout page, SKRect area, ParsedChart chart)
        {
            DrawTitle(canvas, page, area, chart, area.Top + area.Height * 0.10f);
            if (chart.Series.Count == 0)
            {
                return;
            }

            var values = chart.Series[0].Values;
            double sum = 0d;
            foreach (var v in values) if (v.HasValue) sum += Math.Abs(v.Value);
            if (sum <= 0d)
            {
                return;
            }

            if (chart.IsOfPie)
            {
                DrawOfPie(canvas, page, area, chart, values);
                return;
            }

            var size = Math.Min(area.Width, area.Height) * 0.6f;
            var rect = new SKRect(area.MidX - size / 2f, area.Top + area.Height * 0.2f, area.MidX + size / 2f, area.Top + area.Height * 0.2f + size);
            var start = -90f;
            for (var i = 0; i < values.Count; i++)
            {
                if (!values[i].HasValue) continue;
                var sweep = (float)(Math.Abs(values[i].Value) / sum * 360d);
                using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = PiePointColor(chart, i), IsAntialias = true })
                using (var path = new SKPath())
                {
                    path.MoveTo(rect.MidX, rect.MidY);
                    path.ArcTo(rect, start, sweep, false);
                    path.Close();
                    canvas.DrawPath(path, fill);
                }
                start += sweep;
            }

            DrawPlotAreaBorder(canvas, page, area, ResolvePieBorderRect(area, chart, rect), chart);
            DrawLegend(canvas, page, area, rect, chart);
        }

        private void DrawOfPie(SKCanvas canvas, PageLayout page, SKRect area, ParsedChart chart, List<double?> values)
        {
            var secondaryIndices = ResolveOfPieSecondaryIndices(chart, values);
            if (secondaryIndices.Count == 0 || secondaryIndices.Count >= values.Count)
            {
                chart.IsOfPie = false;
                DrawPie(canvas, page, area, chart);
                chart.IsOfPie = true;
                return;
            }

            var plotRect = ResolvePiePlotRect(area, chart);

            var secondaryFlags = new bool[values.Count];
            for (var i = 0; i < secondaryIndices.Count; i++)
            {
                secondaryFlags[secondaryIndices[i]] = true;
            }

            var primaryValues = new List<double>();
            var primaryColors = new List<SKColor>();
            var secondaryValues = new List<double>();
            var secondaryColors = new List<SKColor>();

            double secondarySum = 0d;
            for (var i = 0; i < values.Count; i++)
            {
                if (!values[i].HasValue)
                {
                    continue;
                }

                var absolute = Math.Abs(values[i].Value);
                if (secondaryFlags[i])
                {
                    secondaryValues.Add(absolute);
                    secondaryColors.Add(PiePointColor(chart, i));
                    secondarySum += absolute;
                }
                else
                {
                    primaryValues.Add(absolute);
                    primaryColors.Add(PiePointColor(chart, i));
                }
            }

            if (secondaryValues.Count == 0 || secondarySum <= 0d)
            {
                chart.IsOfPie = false;
                DrawPie(canvas, page, area, chart);
                chart.IsOfPie = true;
                return;
            }

            primaryValues.Add(secondarySum);
            primaryColors.Add(PiePointColor(chart, values.Count));

            var largerChartBonus = ResolveOfPieLargerChartBonus(area);
            var baseMainSize = Math.Min(area.Width, area.Height) * (0.515f + largerChartBonus * 0.02f);
            var maxMainSize = Math.Min(plotRect.Width * (0.64f + largerChartBonus * 0.03f), plotRect.Height * 0.96f);
            var mainSize = Math.Min(baseMainSize, maxMainSize) * 1.1f;
            var secondSize = mainSize * (float)Math.Max(0.35d, Math.Min(1.2d, chart.OfPieSecondPieSizePercent / 100d));
            var gapRatio = (float)Math.Max(0.1d, Math.Min(3d, chart.OfPieGapWidthPercent / 100d));
            var centerGap = 48.965f + 32.643f * gapRatio;
            var groupWidth = mainSize + centerGap + secondSize;
            var groupWidthLimit = plotRect.Width * (0.94f + largerChartBonus * 0.025f);
            if (groupWidth > groupWidthLimit)
            {
                var scale = groupWidthLimit / groupWidth;
                mainSize *= scale;
                secondSize *= scale;
                centerGap *= scale;
                groupWidth = mainSize + centerGap + secondSize;
            }

            var groupHeight = Math.Max(mainSize, secondSize);
            var left = plotRect.MidX - groupWidth / 2f;
            var top = plotRect.MidY - groupHeight / 2f;
            var mainRect = new SKRect(left, top + (mainSize < secondSize ? (secondSize - mainSize) * 0.5f : 0f), left + mainSize, top + (mainSize < secondSize ? (secondSize - mainSize) * 0.5f : 0f) + mainSize);
            var secondLeft = mainRect.Right + centerGap;
            var secondRect = new SKRect(secondLeft, top + (secondSize < mainSize ? (mainSize - secondSize) * 0.5f : 0f), secondLeft + secondSize, top + (secondSize < mainSize ? (mainSize - secondSize) * 0.5f : 0f) + secondSize);

            float aggregateStart;
            float aggregateSweep;
            DrawPieSlices(canvas, mainRect, primaryValues, primaryColors, 70f, out aggregateStart, out aggregateSweep);

            float secondaryStart;
            float secondarySweep;
            DrawPieSlices(canvas, secondRect, secondaryValues, secondaryColors, aggregateStart + 145f, out secondaryStart, out secondarySweep);

            if (chart.HasOfPieSeparatorLines)
            {
                DrawOfPieSeparatorLines(canvas, chart, mainRect, aggregateStart, aggregateSweep, secondRect);
            }

            DrawPlotAreaBorder(canvas, page, area, ResolveOfPieBorderRect(area, chart), chart);
            DrawLegend(canvas, page, area, mainRect, chart);
        }

        private SKRect ResolvePiePlotRect(SKRect area, ParsedChart chart)
        {
            var largerChartBonus = chart != null && chart.IsOfPie ? ResolveOfPieLargerChartBonus(area) : 0f;
            var left = area.Left + area.Width * 0.08f;
            var right = area.Right - area.Width * 0.08f;
            var titleReserve = string.IsNullOrEmpty(chart != null ? chart.Title : null) ? 0.12f : 0.19f;
            var legendReserve = chart != null && chart.LegendPosition != null ? 0.22f : 0.10f;
            if (largerChartBonus > 0f)
            {
                titleReserve -= 0.025f * largerChartBonus;
                legendReserve -= 0.03f * largerChartBonus;
            }

            var top = area.Top + area.Height * titleReserve;
            var bottom = area.Bottom - area.Height * legendReserve;
            var rect = new SKRect(left, top, right, bottom);

            if (largerChartBonus > 0f)
            {
                rect = ExpandRectAboutCenter(rect, 1f + 0.10f * largerChartBonus, 1f + 0.20f * largerChartBonus);
                rect = InsetAndClamp(area, rect.Left, rect.Top, rect.Right, rect.Bottom);
            }

            return rect;
        }

        private SKRect ResolvePieBorderRect(SKRect area, ParsedChart chart, SKRect contentRect)
        {
            var horizontalPadding = area.Width * 0.035f;
            var verticalPadding = area.Height * 0.03f;
            var left = contentRect.Left - horizontalPadding;
            var right = contentRect.Right + horizontalPadding;
            var top = contentRect.Top - verticalPadding;
            var bottom = contentRect.Bottom + verticalPadding;

            var titleBottom = area.Top + (string.IsNullOrEmpty(chart != null ? chart.Title : null) ? area.Height * 0.08f : area.Height * 0.18f);
            if (top < titleBottom)
            {
                top = titleBottom;
            }

            var legendTop = area.Bottom - (chart != null && chart.LegendPosition != null ? area.Height * 0.24f : area.Height * 0.08f);
            if (bottom > legendTop)
            {
                bottom = legendTop;
            }

            return InsetAndClamp(area, left, top, right, bottom);
        }

        private SKRect ResolveOfPieBorderRect(SKRect area, ParsedChart chart)
        {
            var plotRect = ResolvePiePlotRect(area, chart);
            var largerChartBonus = ResolveOfPieLargerChartBonus(area);
            var borderRect = ExpandRectAboutCenter(plotRect, 1f + 0.116f * largerChartBonus, 1.29f);

            var topLimit = area.Top + (string.IsNullOrEmpty(chart != null ? chart.Title : null) ? area.Height * 0.06f : area.Height * 0.12f);
            if (borderRect.Top < topLimit)
            {
                borderRect.Top = topLimit;
            }

            var bottomLimit = area.Bottom - (chart != null && chart.LegendPosition != null ? area.Height * 0.18f : area.Height * 0.05f);
            if (borderRect.Bottom > bottomLimit)
            {
                borderRect.Bottom = bottomLimit;
            }

            return InsetAndClamp(area, borderRect.Left, borderRect.Top, borderRect.Right, borderRect.Bottom);
        }

        private void DrawPieSlices(SKCanvas canvas, SKRect rect, List<double> values, List<SKColor> colors, float initialStart, out float lastStart, out float lastSweep)
        {
            lastStart = initialStart;
            lastSweep = 0f;
            double sum = 0d;
            for (var i = 0; i < values.Count; i++)
            {
                sum += values[i];
            }

            if (sum <= 0d)
            {
                return;
            }

            var start = initialStart;
            for (var i = 0; i < values.Count; i++)
            {
                var sweep = (float)(values[i] / sum * 360d);
                using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = i < colors.Count ? colors[i] : SKColors.Gray, IsAntialias = true })
                using (var path = new SKPath())
                {
                    path.MoveTo(rect.MidX, rect.MidY);
                    path.ArcTo(rect, start, sweep, false);
                    path.Close();
                    canvas.DrawPath(path, fill);
                }

                lastStart = start;
                lastSweep = sweep;
                start += sweep;
            }
        }

        private void DrawOfPieSeparatorLines(SKCanvas canvas, ParsedChart chart, SKRect mainRect, float aggregateStart, float aggregateSweep, SKRect secondRect)
        {
            SKPoint topMain;
            SKPoint topSecond;
            SKPoint bottomMain;
            SKPoint bottomSecond;
            if (!TryResolveOuterTangents(mainRect, secondRect, out topMain, out topSecond, out bottomMain, out bottomSecond))
            {
                var firstAngle = aggregateStart;
                var secondAngle = aggregateStart + aggregateSweep;
                topMain = PieArcPoint(mainRect, firstAngle);
                bottomMain = PieArcPoint(mainRect, secondAngle);
                topSecond = new SKPoint(secondRect.Left, secondRect.MidY - secondRect.Height * 0.22f);
                bottomSecond = new SKPoint(secondRect.Left, secondRect.MidY + secondRect.Height * 0.22f);

                if (topMain.Y > bottomMain.Y)
                {
                    var swap = topMain;
                    topMain = bottomMain;
                    bottomMain = swap;
                }
            }

            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.OfPieSeparatorLineColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.OfPieSeparatorLineWidthPt),
                IsAntialias = true
            })
            {
                canvas.DrawLine(topMain, topSecond, paint);
                canvas.DrawLine(bottomMain, bottomSecond, paint);
            }
        }

        private static bool TryResolveOuterTangents(SKRect leftCircle, SKRect rightCircle, out SKPoint topLeft, out SKPoint topRight, out SKPoint bottomLeft, out SKPoint bottomRight)
        {
            topLeft = SKPoint.Empty;
            topRight = SKPoint.Empty;
            bottomLeft = SKPoint.Empty;
            bottomRight = SKPoint.Empty;

            var x1 = leftCircle.MidX;
            var y1 = leftCircle.MidY;
            var r1 = leftCircle.Width * 0.5f;
            var x2 = rightCircle.MidX;
            var y2 = rightCircle.MidY;
            var r2 = rightCircle.Width * 0.5f;

            var dx = x2 - x1;
            var dy = y2 - y1;
            var distanceSquared = dx * dx + dy * dy;
            if (distanceSquared <= 0.0001f)
            {
                return false;
            }

            var radiusDelta = r1 - r2;
            var hSquared = distanceSquared - radiusDelta * radiusDelta;
            if (hSquared <= 0.0001f)
            {
                return false;
            }

            var inv = 1f / distanceSquared;
            var h = (float)Math.Sqrt(hSquared);
            var nxTop = (dx * radiusDelta - dy * h) * inv;
            var nyTop = (dy * radiusDelta + dx * h) * inv;
            var nxBottom = (dx * radiusDelta + dy * h) * inv;
            var nyBottom = (dy * radiusDelta - dx * h) * inv;

            topLeft = new SKPoint(x1 + r1 * nxTop, y1 + r1 * nyTop);
            topRight = new SKPoint(x2 + r2 * nxTop, y2 + r2 * nyTop);
            bottomLeft = new SKPoint(x1 + r1 * nxBottom, y1 + r1 * nyBottom);
            bottomRight = new SKPoint(x2 + r2 * nxBottom, y2 + r2 * nyBottom);

            if (topLeft.Y > bottomLeft.Y)
            {
                var swapLeft = topLeft;
                topLeft = bottomLeft;
                bottomLeft = swapLeft;

                var swapRight = topRight;
                topRight = bottomRight;
                bottomRight = swapRight;
            }

            return true;
        }

        private static SKPoint PieArcPoint(SKRect rect, float angleDeg)
        {
            var radians = (float)(Math.PI / 180d * angleDeg);
            return new SKPoint(
                rect.MidX + (float)Math.Cos(radians) * rect.Width * 0.5f,
                rect.MidY + (float)Math.Sin(radians) * rect.Height * 0.5f);
        }

        private List<int> ResolveOfPieSecondaryIndices(ParsedChart chart, List<double?> values)
        {
            var result = new List<int>();
            if (chart == null || values == null)
            {
                return result;
            }

            var validCount = 0;
            for (var i = 0; i < values.Count; i++)
            {
                if (values[i].HasValue)
                {
                    validCount++;
                }
            }

            if (validCount <= 1)
            {
                return result;
            }

            var splitType = chart.OfPieSplitType;
            if (string.IsNullOrEmpty(splitType) || string.Equals(splitType, "auto", StringComparison.Ordinal) || string.Equals(splitType, "pos", StringComparison.Ordinal))
            {
                var count = chart.OfPieSplitPosition.HasValue ? (int)Math.Round(chart.OfPieSplitPosition.Value) : 2;
                if (count < 1)
                {
                    count = 1;
                }

                if (count >= validCount)
                {
                    count = validCount - 1;
                }

                for (var i = values.Count - 1; i >= 0 && result.Count < count; i--)
                {
                    if (values[i].HasValue)
                    {
                        result.Insert(0, i);
                    }
                }

                return result;
            }

            if (string.Equals(splitType, "val", StringComparison.Ordinal) && chart.OfPieSplitPosition.HasValue)
            {
                for (var i = 0; i < values.Count; i++)
                {
                    if (values[i].HasValue && Math.Abs(values[i].Value) <= chart.OfPieSplitPosition.Value)
                    {
                        result.Add(i);
                    }
                }

                if (result.Count > 0 && result.Count < validCount)
                {
                    return result;
                }
            }

            if (string.Equals(splitType, "percent", StringComparison.Ordinal) && chart.OfPieSplitPosition.HasValue)
            {
                double sum = 0d;
                for (var i = 0; i < values.Count; i++)
                {
                    if (values[i].HasValue)
                    {
                        sum += Math.Abs(values[i].Value);
                    }
                }

                if (sum > 0d)
                {
                    for (var i = 0; i < values.Count; i++)
                    {
                        if (values[i].HasValue && Math.Abs(values[i].Value) / sum * 100d <= chart.OfPieSplitPosition.Value)
                        {
                            result.Add(i);
                        }
                    }

                    if (result.Count > 0 && result.Count < validCount)
                    {
                        return result;
                    }
                }
            }

            for (var i = values.Count - 1; i >= 0 && result.Count < Math.Min(3, validCount - 1); i--)
            {
                if (values[i].HasValue)
                {
                    result.Insert(0, i);
                }
            }

            return result;
        }

        private SKColor PiePointColor(ParsedChart chart, int pointIndex)
        {
            if (chart != null && chart.PiePointColors != null && pointIndex >= 0 && pointIndex < chart.PiePointColors.Count)
            {
                return chart.PiePointColors[pointIndex];
            }

            return _context.Colors.ResolveSchemeName("accent" + ((pointIndex % 6) + 1), SKColors.Gray);
        }

        // --- axis math ---

        private static int MaxPointCount(ParsedChart chart)
        {
            var n = chart.Categories.Count;
            foreach (var s in chart.Series) n = Math.Max(n, s.Values.Count);
            return n;
        }

        private void DrawDateCategoryAxis(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            double plotMinDate;
            double plotMaxDate;
            if (!ChartDateAxisMath.TryGetDateAxisPlotRange(chart, _dateSystem, out plotMinDate, out plotMaxDate))
            {
                return;
            }

            double minDate;
            double maxDate;
            if (!ChartDateAxisMath.TryGetDateAxisRange(chart, out minDate, out maxDate))
            {
                return;
            }

            var firstTick = ChartDateAxisMath.StartOfMonth(minDate, _dateSystem);
            var lastTick = ChartDateAxisMath.StartOfMonth(maxDate, _dateSystem);
            var layout = ResolveDateAxisLayout(chart, plot.Width);
            var visibleSlice = ResolveVisibleChartSlice(page, area);
            var current = firstTick;
            var rotatedClip = new SKRect(
                Math.Max(plot.Left - 36f, visibleSlice.Left),
                plot.Bottom + 2f,
                Math.Min(plot.Right + 18f, visibleSlice.Right),
                visibleSlice.Bottom);
            while (current <= lastTick + 0.1d)
            {
                var x = ChartDateAxisMath.DateCategoryX(current, plotMinDate, plotMaxDate, plot);
                if (chart.HasCategoryGridlines)
                {
                    if (!IsInteriorVerticalGridline(x, plot))
                    {
                        current = ChartDateAxisMath.AddMonthsSerial(current, layout.MonthStep, _dateSystem);
                        continue;
                    }

                    using (var grid = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        Color = chart.CategoryGridlineColor,
                        StrokeWidth = (float)Math.Max(0.5d, chart.CategoryGridlineWidthPt),
                        IsAntialias = false
                    })
                    {
                        canvas.DrawLine(x, plot.Top, x, plot.Bottom, grid);
                    }
                }

                if (_suppressCategoryAxisLabels)
                {
                    current = ChartDateAxisMath.AddMonthsSerial(current, layout.MonthStep, _dateSystem);
                    continue;
                }

                var label = FormatValue(current, string.IsNullOrEmpty(chart.CategoryFormatCode) ? "mmm-yy" : chart.CategoryFormatCode);
                if (Math.Abs(layout.RotationDeg) > 0.1f)
                {
                    var allowTrailingClipOnLastSlice = IsLastHorizontalChartSlice(page, area);
                    var allowLeadingClipOnContinuationSlice = area.Left < visibleSlice.Left - 0.1f;
                    var leftCull = allowLeadingClipOnContinuationSlice ? visibleSlice.Left - 72f : visibleSlice.Left - 10f;
                    if (x < leftCull || (!allowTrailingClipOnLastSlice && x > visibleSlice.Right + 24f))
                    {
                        current = ChartDateAxisMath.AddMonthsSerial(current, layout.MonthStep, _dateSystem);
                        continue;
                    }

                    DrawRotatedText(canvas, page, rotatedClip, label, x + 2f, plot.Bottom + 6f, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Right, VAlign.Top, layout.RotationDeg);
                }
                else
                {
                    DrawText(canvas, page, area, label, x, plot.Bottom + 4f, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Center, VAlign.Top);
                }
                current = ChartDateAxisMath.AddMonthsSerial(current, layout.MonthStep, _dateSystem);
            }
        }

        private void DrawBarValueAxis(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart, double min, double max, double major)
        {
            if (max <= min || major <= 0d)
            {
                return;
            }

            major = ResolveRenderedMajorUnit(chart, plot, major);

            using (var grid = new SKPaint { Style = SKPaintStyle.Stroke, Color = GridColor, StrokeWidth = 1f, IsAntialias = false })
            {
                var steps = (int)Math.Round((max - min) / major);
                for (var i = 0; i <= steps; i++)
                {
                    var value = min + i * major;
                    var x = ValueToX(value, min, max, plot);

                    if (chart.HasValueGridlines)
                    {
                        canvas.DrawLine(x, plot.Top, x, plot.Bottom, grid);
                    }

                    var label = FormatValue(value, chart.ValueFormatCode);
                    var width = MeasureText(label, LabelSizePt);
                    var clip = new SKRect(x - width * 0.5f - 2f, plot.Bottom + 2f, x + width * 0.5f + 2f, plot.Bottom + LabelSizePt * 2.2f);
                    DrawText(canvas, page, clip, label, x, plot.Bottom + 4f, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Center, VAlign.Top);
                }
            }
        }

        private void DrawBarCategoryAxis(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            using (var axis = new SKPaint { Style = SKPaintStyle.Stroke, Color = GridColor, StrokeWidth = 1f, IsAntialias = false })
            {
                canvas.DrawLine(plot.Left, plot.Top, plot.Left, plot.Bottom, axis);
            }

            if (_suppressCategoryAxisLabels)
            {
                return;
            }

            var count = chart.Categories.Count;
            if (count == 0)
            {
                count = chart.CategoryValues.Count;
            }

            if (count == 0)
            {
                return;
            }

            for (var i = 0; i < count; i++)
            {
                if (chart.HasCategoryGridlines)
                {
                    var gridlineY = CategoryGridlineY(i, count, plot, chart.CrossBetween, ShouldReverseBarCategoryAxis(chart));
                    using (var grid = new SKPaint
                    {
                        Style = SKPaintStyle.Stroke,
                        Color = chart.CategoryGridlineColor,
                        StrokeWidth = (float)Math.Max(0.5d, chart.CategoryGridlineWidthPt),
                        IsAntialias = false
                    })
                    {
                        canvas.DrawLine(plot.Left, gridlineY, plot.Right, gridlineY, grid);
                    }
                }

                var y = CategoryY(i, count, plot, chart.CrossBetween, ShouldReverseBarCategoryAxis(chart));
                var label = ResolveCategoryLabel(chart, i);
                if (string.IsNullOrEmpty(label))
                {
                    continue;
                }

                var width = MeasureText(label, LabelSizePt);
                var clip = new SKRect(plot.Left - width - 8f, y - LabelSizePt, plot.Left - 2f, y + LabelSizePt);
                DrawText(canvas, page, clip, label, plot.Left - 5f, y, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Right, VAlign.Middle);
            }
        }

        private float PointX(ParsedChart chart, int index, int count, SKRect plot)
        {
            if (chart != null && chart.HasDateCategoryAxis)
            {
                double plotMinDate;
                double plotMaxDate;
                if (ChartDateAxisMath.TryGetDateAxisPlotRange(chart, _dateSystem, out plotMinDate, out plotMaxDate)
                    && index >= 0
                    && index < chart.CategoryValues.Count
                    && chart.CategoryValues[index].HasValue)
                {
                    var serial = ChartDateAxisMath.NormalizePointSerial(chart, chart.CategoryValues[index].Value, _dateSystem);
                    return ChartDateAxisMath.DateCategoryX(serial, plotMinDate, plotMaxDate, plot);
                }
            }

            return CategoryX(index, count, plot, chart != null && chart.CrossBetween);
        }

        private float CategoryY(int index, int count, SKRect plot, bool crossBetween)
        {
            return CategoryY(index, count, plot, crossBetween, false);
        }

        private float CategoryY(int index, int count, SKRect plot, bool crossBetween, bool reverse)
        {
            if (count <= 0)
            {
                return plot.MidY;
            }

            var renderIndex = reverse ? count - 1 - index : index;

            if (crossBetween)
            {
                return plot.Top + (renderIndex + 0.5f) / count * plot.Height;
            }

            if (count == 1)
            {
                return plot.MidY;
            }

            return plot.Top + (float)renderIndex / (count - 1) * plot.Height;
        }

        private float CategoryGridlineY(int index, int count, SKRect plot, bool crossBetween, bool reverse)
        {
            if (count <= 1)
            {
                return plot.MidY;
            }

            if (crossBetween)
            {
                var renderIndex = reverse ? count - 1 - index : index;
                if (renderIndex >= count - 1)
                {
                    renderIndex = count - 2;
                }

                return plot.Top + (renderIndex + 1f) / count * plot.Height;
            }

            return CategoryY(index, count, plot, false, reverse);
        }

        private int ResolveBarCategoryIndex(ParsedChart chart, int index, int count)
        {
            if (!ShouldReverseBarCategoryAxis(chart))
            {
                return index;
            }

            return count - 1 - index;
        }

        private bool ShouldReverseBarCategoryAxis(ParsedChart chart)
        {
            if (chart == null || chart.Kind != ChartKind.Bar)
            {
                return false;
            }

            return chart.CategoryAxisMinMax;
        }

        private string ResolveCategoryLabel(ParsedChart chart, int index)
        {
            if (chart == null || index < 0)
            {
                return string.Empty;
            }

            if (index < chart.Categories.Count && !string.IsNullOrEmpty(chart.Categories[index]))
            {
                return chart.Categories[index];
            }

            if (index < chart.CategoryValues.Count && chart.CategoryValues[index].HasValue)
            {
                return FormatValue(chart.CategoryValues[index].Value, string.IsNullOrEmpty(chart.CategoryFormatCode) ? "mmm-yy" : chart.CategoryFormatCode);
            }

            return string.Empty;
        }

        private static float CategoryX(int index, int count, SKRect plot, bool crossBetween)
        {
            if (count <= 0)
            {
                return plot.MidX;
            }

            if (crossBetween)
            {
                // Points sit at the center of each category band (Excel's default), inset from the
                // plot edges by half a band.
                return plot.Left + (index + 0.5f) / count * plot.Width;
            }

            if (count == 1)
            {
                return plot.MidX;
            }

            // Points sit on the axis ticks: first at the left edge, last at the right edge.
            return plot.Left + (float)index / (count - 1) * plot.Width;
        }

        private static float ValueToY(double value, double min, double max, SKRect plot)
        {
            var t = (value - min) / (max - min);
            return plot.Bottom - (float)t * plot.Height;
        }

        private static float ValueToX(double value, double min, double max, SKRect plot)
        {
            var t = (value - min) / (max - min);
            return plot.Left + (float)t * plot.Width;
        }

        private static int FirstValueIndex(List<double?> values)
        {
            for (var i = 0; i < values.Count; i++) if (values[i].HasValue) return i;
            return 0;
        }

        private static int LastValueIndex(List<double?> values)
        {
            for (var i = values.Count - 1; i >= 0; i--) if (values[i].HasValue) return i;
            return 0;
        }

        private SKRect ResolveVisibleChartSlice(PageLayout page, SKRect area)
        {
            if (page == null || page.Sheet == null)
            {
                return area;
            }

            var contentWidth = (float)(page.Sheet.ColumnStartPt[page.EndColumn + 1] - page.Sheet.ColumnStartPt[page.StartColumn]);
            var contentHeight = (float)(page.Sheet.RowStartPt[page.EndRow + 1] - page.Sheet.RowStartPt[page.StartRow]);
            var left = Math.Max(area.Left, 0f);
            var top = Math.Max(area.Top, 0f);
            var right = Math.Min(area.Right, contentWidth);
            var bottom = Math.Min(area.Bottom, contentHeight);
            return new SKRect(left, top, Math.Max(left, right), Math.Max(top, bottom));
        }

        private static void ComputeValueAxis(ParsedChart chart, out double min, out double max, out double major)
        {
            double dmin;
            double dmax;
            ChartStackingMath.ComputeDataRange(chart, out dmin, out dmax);

            if (dmin == double.MaxValue)
            {
                min = 0d; max = 1d; major = 0.2d;
                return;
            }

            if (chart != null && chart.IsPercentStacked)
            {
                min = chart.AxisMin ?? 0d;
                max = chart.AxisMax ?? 1d;
                major = chart.MajorUnit ?? 0.2d;
                if (major <= 0d)
                {
                    major = 0.2d;
                }

                if (max <= min)
                {
                    max = min + major;
                }

                return;
            }

            min = chart.AxisMin ?? (dmin >= 0d ? 0d : -NiceNumber(-dmin));
            major = chart.MajorUnit ?? NiceNumber((dmax - min) / 8d);
            if (major <= 0d) major = 1d;

            if (chart.AxisMax.HasValue)
            {
                max = chart.AxisMax.Value;
            }
            else
            {
                max = Math.Ceiling((dmax - min) / major) * major + min;
                // Radar charts use the rounded outer ring directly instead of adding
                // an extra headroom band above the data maximum.
                if ((chart == null || chart.Kind != ChartKind.Radar) && max - dmax < 0.15d * major) max += major;
            }

            if (max <= min) max = min + major;
        }

        private static double NiceNumber(double value)
        {
            if (value <= 0d) return 1d;
            var exponent = Math.Floor(Math.Log10(value));
            var fraction = value / Math.Pow(10d, exponent);
            double nice;
            if (fraction < 1.5d) nice = 1d;
            else if (fraction < 3d) nice = 2d;
            else if (fraction < 7d) nice = 5d;
            else nice = 10d;
            return nice * Math.Pow(10d, exponent);
        }

        private string FormatValue(double value, string formatCode)
        {
            var text = ChartXmlParser.FormatLabel(value, formatCode, _dateSystem, _chartCulture);
            return text != null ? text.TrimEnd() : string.Empty;
        }

        private float MeasureMaxValueAxisLabelWidth(ParsedChart chart)
        {
            double min;
            double max;
            double major;
            ComputeValueAxis(chart, out min, out max, out major);
            if (max <= min || major <= 0d)
            {
                return 0f;
            }

            var maxWidth = 0f;
            var steps = (int)Math.Round((max - min) / major);
            for (var i = 0; i <= steps; i++)
            {
                var value = min + i * major;
                var label = FormatValue(value, chart.ValueFormatCode);
                var width = MeasureText(label, LabelSizePt);
                if (width > maxWidth)
                {
                    maxWidth = width;
                }
            }

            return maxWidth;
        }

        private static SKRect InsetAndClamp(SKRect area, float left, float top, float right, float bottom)
        {
            if (left < area.Left)
            {
                left = area.Left;
            }

            if (top < area.Top)
            {
                top = area.Top;
            }

            if (right > area.Right)
            {
                right = area.Right;
            }

            if (bottom > area.Bottom)
            {
                bottom = area.Bottom;
            }

            if (right <= left)
            {
                left = area.Left;
                right = area.Right;
            }

            if (bottom <= top)
            {
                top = area.Top;
                bottom = area.Bottom;
            }

            return new SKRect(left, top, right, bottom);
        }

        private ChartDateAxisLayout ResolveDateAxisLayout(ParsedChart chart, float availableWidth)
        {
            return ChartDateAxisLayout.Resolve(chart, _dateSystem, availableWidth);
        }

        private void DrawPlotAreaBorder(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            if (canvas == null || chart == null || !chart.HasPlotAreaBorder)
            {
                return;
            }

            var strokeWidth = (float)Math.Max(0.5d, chart.PlotAreaBorderWidthPt);
            if (chart.Kind == ChartKind.Pie)
            {
                strokeWidth = Math.Max(0.5f, strokeWidth * 0.75f);
            }

            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.PlotAreaBorderColor,
                StrokeWidth = strokeWidth,
                IsAntialias = false
            })
            {
                if (chart.Kind == ChartKind.Pie)
                {
                    canvas.DrawLine(plot.Left, plot.Top, plot.Right, plot.Top, paint);
                    canvas.DrawLine(plot.Left, plot.Bottom, plot.Right, plot.Bottom, paint);
                    canvas.DrawLine(plot.Left, plot.Top, plot.Left, plot.Bottom, paint);
                    canvas.DrawLine(plot.Right, plot.Top, plot.Right, plot.Bottom, paint);
                    return;
                }

                var visibleSlice = ResolveVisibleChartSlice(page, area);
                var visiblePlot = IntersectRect(plot, visibleSlice);
                var drawRightEdge = IsLastHorizontalChartSlice(page, area);
                var verticalStrokeWidth = Math.Max(0.5f, paint.StrokeWidth * 0.7f);
                canvas.DrawLine(plot.Left, plot.Top, plot.Right, plot.Top, paint);
                paint.StrokeWidth = verticalStrokeWidth;
                canvas.DrawLine(visiblePlot.Left, plot.Top, visiblePlot.Left, plot.Bottom, paint);
                if (drawRightEdge)
                {
                    canvas.DrawLine(visiblePlot.Right, plot.Top, visiblePlot.Right, plot.Bottom, paint);
                }
            }
        }

        private static SKRect UnionRect(SKRect a, SKRect b)
        {
            return new SKRect(
                Math.Min(a.Left, b.Left),
                Math.Min(a.Top, b.Top),
                Math.Max(a.Right, b.Right),
                Math.Max(a.Bottom, b.Bottom));
        }

        private static SKRect ExpandRectAboutCenter(SKRect rect, float scaleX, float scaleY)
        {
            var halfWidth = rect.Width * scaleX * 0.5f;
            var halfHeight = rect.Height * scaleY * 0.5f;
            return new SKRect(
                rect.MidX - halfWidth,
                rect.MidY - halfHeight,
                rect.MidX + halfWidth,
                rect.MidY + halfHeight);
        }

        private float ResolveOfPieLargerChartBonus(SKRect area)
        {
            var minDimension = Math.Min(area.Width, area.Height);
            if (minDimension <= 170f)
            {
                return 0f;
            }

            if (minDimension >= 230f)
            {
                return 1f;
            }

            return (minDimension - 170f) / 60f;
        }

        private static bool IsInteriorVerticalGridline(float x, SKRect plot)
        {
            return x > plot.Left + 0.5f && x < plot.Right - 0.5f;
        }

        private bool IsLastHorizontalChartSlice(PageLayout page, SKRect area)
        {
            if (page == null || page.Sheet == null)
            {
                return true;
            }

            var visibleSlice = ResolveVisibleChartSlice(page, area);
            return area.Right <= visibleSlice.Right + 0.1f;
        }

        private static SKRect IntersectRect(SKRect a, SKRect b)
        {
            var left = Math.Max(a.Left, b.Left);
            var top = Math.Max(a.Top, b.Top);
            var right = Math.Min(a.Right, b.Right);
            var bottom = Math.Min(a.Bottom, b.Bottom);
            if (right < left)
            {
                right = left;
            }

            if (bottom < top)
            {
                bottom = top;
            }

            return new SKRect(left, top, right, bottom);
        }

        private static double ResolveRenderedMajorUnit(ParsedChart chart, SKRect plot, double major)
        {
            if (chart == null)
            {
                return major;
            }

            if (chart.IsPercentStacked && !chart.MajorUnit.HasValue && major >= 0.1d)
            {
                return 0.1d;
            }

            return major;
        }

        private static double ResolveRenderedMinorUnit(ParsedChart chart, double major)
        {
            if (chart == null)
            {
                return 0d;
            }

            if (chart.MinorUnit.HasValue && chart.MinorUnit.Value > 0d)
            {
                return chart.MinorUnit.Value;
            }

            if (major <= 0d)
            {
                return 0d;
            }

            return major / 5d;
        }

        private float ResolveValueGridlineRight(SKRect plot, ParsedChart chart)
        {
            if (chart != null
                && chart.Kind == ChartKind.Line
                && !chart.IsStacked
                && !chart.IsPercentStacked
                && chart.HasDateCategoryAxis
                && chart.LegendPosition != null)
            {
                return plot.Right + StandardDateAxisGridlineRightExtensionPt;
            }

            if (chart != null
                && chart.Kind == ChartKind.Line
                && chart.IsStacked
                && !chart.IsPercentStacked
                && chart.HasDateCategoryAxis
                && chart.LegendPosition != null)
            {
                return plot.Right + StackedDateAxisGridlineRightExtensionPt;
            }

            return plot.Right;
        }

        // --- text helpers ---

        private enum TextAlign { Left, Center, Right }
        private enum VAlign { Top, Middle, Baseline }

        private float MeasureText(string text, float sizePt)
        {
            return MeasureText(text, sizePt, false);
        }

        private float MeasureText(string text, float sizePt, bool bold)
        {
            return MeasureText(text, ChartFont(sizePt, bold));
        }

        private float MeasureText(string text, FontValue font)
        {
            var fontContext = _context.GetFontContext(font);
            return fontContext.Measure(text ?? string.Empty);
        }

        private void DrawText(SKCanvas canvas, PageLayout page, SKRect clipRect, string text, float x, float y, float sizePt, bool bold, SKColor color, TextAlign hAlign, VAlign vAlign)
        {
            DrawText(canvas, page, clipRect, text, x, y, sizePt, bold, false, color, hAlign, vAlign);
        }

        private void DrawText(SKCanvas canvas, PageLayout page, SKRect clipRect, string text, float x, float y, float sizePt, bool bold, bool italic, SKColor color, TextAlign hAlign, VAlign vAlign)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var fontValue = ChartFont(sizePt, bold, italic);
            var fontContext = _context.GetFontContext(fontValue);
            using (var paint = TextPaint(fontValue, color))
            {
                var width = fontContext.Measure(text);
                var drawX = hAlign == TextAlign.Left ? x : hAlign == TextAlign.Center ? x - width / 2f : x - width;

                var metrics = paint.FontMetrics;
                float baseline;
                if (vAlign == VAlign.Middle) baseline = y - (metrics.Ascent + metrics.Descent) / 2f;
                else if (vAlign == VAlign.Top) baseline = y - metrics.Ascent;
                else baseline = y;

                if (!TryRecordRuns(page, clipRect, fontContext, text, drawX, baseline, paint.Color))
                {
                    DrawRuns(canvas, fontContext, text, drawX, baseline, paint);
                }
            }
        }

        private void DrawTextWithoutPdfOptimization(SKCanvas canvas, SKRect clipRect, string text, float x, float y, FontValue fontValue, SKColor color, TextAlign hAlign, VAlign vAlign)
        {
            if (canvas == null || string.IsNullOrEmpty(text) || fontValue == null)
            {
                return;
            }

            var fontContext = _context.GetFontContext(fontValue);
            using (var paint = TextPaint(fontValue, color))
            {
                var width = fontContext.Measure(text);
                var drawX = hAlign == TextAlign.Left ? x : hAlign == TextAlign.Center ? x - width / 2f : x - width;

                var metrics = paint.FontMetrics;
                float baseline;
                if (vAlign == VAlign.Middle) baseline = y - (metrics.Ascent + metrics.Descent) / 2f;
                else if (vAlign == VAlign.Top) baseline = y - metrics.Ascent;
                else baseline = y;

                canvas.Save();
                canvas.ClipRect(clipRect);
                DrawRuns(canvas, fontContext, text, drawX, baseline, paint);
                canvas.Restore();
            }
        }

        private void DrawRotatedText(SKCanvas canvas, PageLayout page, SKRect clipRect, string text, float x, float y, float sizePt, bool bold, SKColor color, TextAlign hAlign, VAlign vAlign, float angleDegrees)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var fontValue = ChartFont(sizePt, bold);
            var fontContext = _context.GetFontContext(fontValue);
            using (var paint = TextPaint(fontValue, color))
            {
                var width = fontContext.Measure(text);
                var drawX = hAlign == TextAlign.Left ? 0f : hAlign == TextAlign.Center ? -width / 2f : -width;

                var metrics = paint.FontMetrics;
                float baseline;
                if (vAlign == VAlign.Middle) baseline = -(metrics.Ascent + metrics.Descent) / 2f;
                else if (vAlign == VAlign.Top) baseline = -metrics.Ascent;
                else baseline = 0f;

                if (!TryRecordRotatedRuns(page, clipRect, fontContext, text, x, y, drawX, baseline, angleDegrees, paint.Color))
                {
                    canvas.Save();
                    canvas.ClipRect(clipRect);
                    canvas.Translate(x, y);
                    canvas.RotateDegrees(angleDegrees);
                    DrawRuns(canvas, fontContext, text, drawX, baseline, paint);
                    canvas.Restore();
                }
            }
        }

        private bool TryRecordRuns(PageLayout page, SKRect clipRect, CellFontContext fontContext, string text, float x, float baseline, SKColor color)
        {
            var session = _context.PdfTextSession;
            if (session == null || page == null || string.IsNullOrEmpty(text))
            {
                return false;
            }

            // For charts, only hand off ordinary horizontal text to the compact Type3 text writer.
            // Rotated category labels stay on the path-rendering route until the writer can place
            // rotated chart text without inheriting Skia's page transform.
            if (!IsSimpleChartTextRun(text, clipRect))
            {
                return false;
            }

            var runs = fontContext.SplitRuns(text);
            var cursor = x;
            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                session.RecordChartRun(page, clipRect, run.Text, run.Typeface, color, fontContext.SizePt, cursor, baseline);
                cursor += run.WidthPt;
            }

            return true;
        }

        private bool TryRecordRotatedRuns(PageLayout page, SKRect clipRect, CellFontContext fontContext, string text, float originX, float originY, float x, float baseline, float rotationDeg, SKColor color)
        {
            var session = _context.PdfTextSession;
            if (session == null || page == null || string.IsNullOrEmpty(text))
            {
                return false;
            }

            if (!IsTokenizableChartText(text))
            {
                return false;
            }

            var runs = fontContext.SplitRuns(text);
            var cursor = x;
            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                session.RecordChartRotatedRun(page, clipRect, run.Text, run.Typeface, color, fontContext.SizePt, originX, originY, cursor, baseline, rotationDeg);
                cursor += run.WidthPt;
            }

            return true;
        }

        private static bool IsTokenizableChartText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                if (ch == '\r' || ch == '\n')
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsSimpleChartTextRun(string text, SKRect clipRect)
        {
            if (clipRect.Height > 32f)
            {
                return false;
            }

            return IsTokenizableChartText(text);
        }

        private static void DrawRuns(SKCanvas canvas, CellFontContext fontContext, string text, float x, float baseline, SKPaint paint)
        {
            var runs = fontContext.SplitRuns(text);
            var cursor = x;

            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                using (var font = new SKFont(run.Typeface, fontContext.SizePt))
                {
                    font.Subpixel = true;
                    PdfTextPathRenderer.DrawText(canvas, run.Text, cursor, baseline, font, paint.Color);
                }

                cursor += run.WidthPt;
            }
        }

        private SKPaint TextPaint(FontValue font, SKColor color)
        {
            return new SKPaint
            {
                Color = color,
                IsAntialias = true,
                Typeface = _context.Fonts.Resolve(font),
                TextSize = (float)font.Size,
                TextAlign = SKTextAlign.Left,
                SubpixelText = true,
            };
        }

        private FontValue ResolveTitleFont(ParsedChart chart)
        {
            if (chart != null && chart.TitleFont != null)
            {
                return chart.TitleFont;
            }

            return ChartFont(14f, true, false);
        }

        private FontValue ChartFont(float sizePt, bool bold = false)
        {
            return ChartFont(sizePt, bold, false);
        }

        private FontValue ChartFont(float sizePt, bool bold, bool italic)
        {
            return new FontValue { Name = "Calibri", Size = sizePt, Bold = bold, Italic = italic };
        }
    }
}
