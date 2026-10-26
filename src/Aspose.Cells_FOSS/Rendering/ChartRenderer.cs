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
        private float LabelSizePt = 9f;
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
            LabelSizePt = (float)(chart.TextSizePt ?? 9d);
            _suppressCategoryAxisLabels = suppressCategoryAxisLabels;

            DrawFrame(canvas, area, chart);

            if (page != null && chart.IsArea3D && (area.Left < 0f || area.Top < 0f))
            {
                return;
            }

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

            if (chart.Kind == ChartKind.Scatter && chart.ScatterVaryColors
                && chart.ScatterCategoryAxisDeleted && chart.ScatterValueAxisDeleted
                && chart.ScatterTitleOverlay && chart.RoundedCorners)
            {
                DrawHiddenAxisScatter(canvas, page, area, chart);
                return;
            }

            double min, max, major;
            ComputeValueAxis(chart, out min, out max, out major);

            var titleCenterY = area.Top + (seriesPlot.Top - area.Top) * 0.5f;
            if (chart.Kind == ChartKind.Scatter && !chart.HasManualTitle)
            {
                titleCenterY = area.Top + (seriesPlot.Top - area.Top)
                    * (chart.LegendPosition == null ? 0.45f : 0.38f);
            }
            if (chart.HasDataTable)
            {
                titleCenterY = area.Top + area.Height * 0.08f;
            }
            else if (chart.Kind == ChartKind.Column && chart.IsBar3D)
            {
                titleCenterY = area.Top + area.Height * 0.08f;
            }
            else if (chart.Kind == ChartKind.Column && chart.IsPercentStacked && !chart.IsBar3D)
            {
                titleCenterY = area.Top + area.Height * 0.06f;
            }
            if (chart.Kind == ChartKind.Radar)
            {
                titleCenterY = area.Top + (seriesPlot.Top - area.Top) * 0.55f;
            }
            else if (chart.Kind == ChartKind.Area && chart.IsArea3D && chart.IsStacked
                && string.Equals(chart.LegendPosition, "t", StringComparison.Ordinal))
            {
                titleCenterY = area.Top + area.Height * 0.09f;
            }

            DrawTitle(canvas, page, area, chart, LineChartDecorationLayout.TitleCenterY(area, chart, titleCenterY));
            if (chart.Kind == ChartKind.Scatter)
            {
                DrawValueAxis(canvas, page, area, seriesPlot, chart, min, max, major);
                DrawScatterCategoryAxis(canvas, page, seriesPlot, chart, min, max);
                if (chart.IsBubble)
                {
                    BubbleChartMath.DrawSeries(canvas, seriesPlot, chart, min, max);
                }
                else
                {
                    DrawScatterSeries(canvas, seriesPlot, chart, min, max);
                }
                DrawPlotAreaBorder(canvas, page, area, seriesPlot, chart);
                DrawAxisTitles(canvas, page, area, seriesPlot, chart);
                if (!string.IsNullOrEmpty(chart.ValueDisplayUnitLabel))
                {
                    var unitX = seriesPlot.Left - 44f;
                    var unitY = seriesPlot.Top + seriesPlot.Height * 0.2f;
                    if (chart.HasManualValueDisplayUnitLabel)
                    {
                        unitX = area.Left + (float)(chart.ValueDisplayUnitLabelX * area.Width) + 8f;
                        unitY = area.Top + (float)(chart.ValueDisplayUnitLabelY * area.Height)
                            + MeasureText(chart.ValueDisplayUnitLabel, 10f) * 0.5f + 3.5f;
                    }
                    DrawRotatedText(canvas, page, area, chart.ValueDisplayUnitLabel,
                        unitX, unitY, 10f, false,
                        chart.AxisLabelColor, TextAlign.Center, VAlign.Middle, -90f);
                }
                DrawLegend(canvas, page, area, seriesPlot, chart);
                return;
            }

            if (chart.Kind == ChartKind.Radar)
            {
                DrawRadar(canvas, page, area, seriesPlot, chart, min, max, major);
                DrawPlotAreaBorder(canvas, page, area, seriesPlot, chart);
                DrawLegend(canvas, page, area, seriesPlot, chart);
                return;
            }

            if (chart.Kind == ChartKind.Bar)
            {
                if (chart.IsBar3D)
                {
                    var bar3DRenderer = new Bar3DChartRenderer(_context, _dateSystem, _chartCulture);
                    bar3DRenderer.Draw(canvas, area, seriesPlot, chart, min, max, major);
                    DrawLegend(canvas, page, area, seriesPlot, chart);
                    return;
                }

                DrawBarValueAxis(canvas, page, area, seriesPlot, chart, min, max, major);
                DrawSeries(canvas, page, seriesPlot, chart, min, max);
                DrawBarCategoryAxis(canvas, page, area, seriesPlot, chart);
                DrawPlotAreaBorder(canvas, page, area, seriesPlot, chart);
                DrawLegend(canvas, page, area, seriesPlot, chart);
                return;
            }
            else if (chart.Kind == ChartKind.Column && chart.IsBar3D)
            {
                var column3DRenderer = new Column3DChartRenderer(_context, _dateSystem, _chartCulture);
                column3DRenderer.Draw(canvas, area, chart, min, max, major);
                DrawLegend(canvas, page, area, seriesPlot, chart);
                return;
            }
            else if (chart.Kind == ChartKind.Area && chart.IsArea3D)
            {
                var area3DRenderer = new Area3DChartRenderer(_context, _dateSystem, _chartCulture);
                area3DRenderer.Draw(canvas, area, seriesPlot, chart, min, max, major);
                DrawLegend(canvas, page, area, seriesPlot, chart);
                return;
            }
            else if (chart.Kind == ChartKind.Line && chart.IsLine3D)
            {
                var line3DRenderer = new Line3DChartRenderer(_context, _dateSystem, _chartCulture);
                line3DRenderer.Draw(canvas, page, area, seriesPlot, chart, min, max, major);
                DrawLegend(canvas, page, area, seriesPlot, chart);
                return;
            }
            else
            {
                DrawValueAxis(canvas, page, area, seriesPlot, chart, min, max, major);
                DrawCategoryAxis(canvas, page, area, seriesPlot, chart);
                DrawAxisTitles(canvas, page, area, seriesPlot, chart);
            }
            DrawSeries(canvas, page, seriesPlot, chart, min, max);
            DrawPlotAreaBorder(canvas, page, area, seriesPlot, chart);
            DrawDataTable(canvas, page, area, seriesPlot, chart);
            DrawLegend(canvas, page, area, seriesPlot, chart);
        }

        public void DrawFrame(SKCanvas canvas, SKRect area, ParsedChart chart)
        {
            if (canvas == null)
            {
                return;
            }

            if (chart != null && chart.HasChartAreaGlow)
            {
                var sigma = (float)(chart.ChartAreaGlowRadiusPt / 4d);
                var spread = (float)(chart.ChartAreaGlowRadiusPt / 3d);
                var source = chart.ChartAreaGlowColor;
                var opacity = (byte)Math.Min(255, (int)Math.Round(source.Alpha * 1.25d));
                using (var glow = new SKPaint { Style = SKPaintStyle.Fill,
                    Color = new SKColor(source.Red, source.Green, source.Blue, opacity), IsAntialias = true,
                    MaskFilter = SKMaskFilter.CreateBlur(SKBlurStyle.Normal, sigma) })
                {
                    var glowRect = SKRect.Create(area.Left - spread, area.Top - spread,
                        area.Width + spread * 2f, area.Height + spread * 2f);
                    if (chart.ChartAreaSoftEdgeRadiusPt > 0d)
                    {
                        var cornerRadius = (float)Math.Min(chart.ChartAreaSoftEdgeRadiusPt,
                            Math.Min(glowRect.Width, glowRect.Height) * 0.5d);
                        canvas.DrawRoundRect(glowRect, cornerRadius, cornerRadius, glow);
                    }
                    else
                    {
                        canvas.DrawRect(glowRect, glow);
                    }
                }
            }

            if (chart != null && chart.HasChartAreaGradient
                && chart.ChartAreaGradientColors.Count >= 2
                && chart.ChartAreaGradientColors.Count == chart.ChartAreaGradientPositions.Count)
            {
                var radians = chart.ChartAreaGradientAngleDeg * Math.PI / 180d;
                var xOffset = (float)(Math.Cos(radians) * area.Width * 0.5d);
                var yOffset = (float)(Math.Sin(radians) * area.Height * 0.5d);
                var start = new SKPoint(area.MidX - xOffset, area.MidY - yOffset);
                var end = new SKPoint(area.MidX + xOffset, area.MidY + yOffset);
                using (var shader = SKShader.CreateLinearGradient(
                    start,
                    end,
                    chart.ChartAreaGradientColors.ToArray(),
                    chart.ChartAreaGradientPositions.ToArray(),
                    SKShaderTileMode.Clamp))
                using (var gradientPaint = new SKPaint { Style = SKPaintStyle.Fill, IsAntialias = false, Shader = shader })
                {
                    canvas.DrawRect(area, gradientPaint);
                }
            }
            else
            {
                if (!ChartPatternRenderer.DrawChartArea(canvas, area, chart))
                {
                    var background = SKColors.White;
                    if (chart != null && chart.HasChartAreaFill)
                    {
                        background = chart.ChartAreaFillColor;
                    }

                    if (chart != null && chart.ChartAreaSoftEdgeRadiusPt > 0d)
                    {
                        ChartSoftEdgeRenderer.DrawSolidFill(canvas, area, background, chart.ChartAreaSoftEdgeRadiusPt);
                    }
                    else
                    {
                        using (var paint = new SKPaint { Style = SKPaintStyle.Fill, Color = background, IsAntialias = false })
                        {
                            canvas.DrawRect(area, paint);
                        }
                    }
                }
            }

            var borderColor = chart != null && chart.HasChartAreaBorder ? chart.ChartAreaBorderColor
                : chart != null && chart.RoundedCorners ? new SKColor(0x88, 0x88, 0x88) : BorderColor;
            var borderWidth = chart != null && chart.HasChartAreaBorder
                ? (float)Math.Max(0.5d, chart.ChartAreaBorderWidthPt) : 1f;
            using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, Color = borderColor,
                StrokeWidth = borderWidth, IsAntialias = false })
            {
                if (chart != null && chart.RoundedCorners)
                {
                    canvas.DrawRoundRect(area, 10f, 10f, paint);
                }
                else
                {
                    canvas.DrawRect(area, paint);
                }
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
            if (chart.Kind == ChartKind.Scatter)
            {
                return ScatterChartPlotLayout.Resolve(area, chart);
            }

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

            var isPercentStackedArea = chart.Kind == ChartKind.Area && chart.IsPercentStacked;
            var axisLabelMaxWidth = MeasureMaxValueAxisLabelWidth(chart);
            var leftPadding = Math.Max(area.Width * 0.06f, axisLabelMaxWidth + 14f);
            var left = area.Left + leftPadding;
            var top = area.Top + (chart.Title != null ? area.Height * (isPercentStackedArea ? 0.136f : 0.205f) : area.Height * 0.085f);
            var right = area.Right - area.Width * 0.07f;
            var bottomReserve = isPercentStackedArea && chart.LegendPosition != null ? 0.17f : BottomReserveFraction(chart, area.Width);

            if (chart.Kind == ChartKind.Column && !chart.IsPercentStacked
                && !chart.IsBar3D && chart.Title != null && chart.LegendPosition != null)
            {
                left = area.Left + Math.Max(area.Width * 0.06f, axisLabelMaxWidth + 14.5f);
                top = area.Top + area.Height * 0.185f;
                right = area.Right - area.Width * 0.031f;
                bottomReserve = 0.239f;
            }

            if (chart.Kind == ChartKind.Column && chart.HasDataTable)
            {
                left = area.Left + area.Width * 0.155f;
                top = area.Top + area.Height * 0.187f;
                right = area.Right - area.Width * 0.03f;
                bottomReserve = 0.575f;
            }

            if (chart.Kind == ChartKind.Column && chart.IsPercentStacked
                && !chart.IsBar3D && !chart.HasDataTable)
            {
                top = area.Top + area.Height * 0.136f;
                right = area.Right - area.Width * 0.02f;
                bottomReserve = 0.17f;
            }

            if (chart.Kind == ChartKind.Column && !chart.IsBar3D
                && !string.IsNullOrEmpty(chart.CategoryAxisTitle)
                && !string.IsNullOrEmpty(chart.ValueAxisTitle))
            {
                left = area.Left + Math.Max(area.Width * 0.06f, axisLabelMaxWidth + 14.5f) + 22f;
                top = area.Top + area.Height * 0.172f;
                right = area.Right - area.Width * 0.0285f;
                bottomReserve = 0.324f;
            }

            if (chart.Kind == ChartKind.Area && chart.IsArea3D && chart.IsStacked
                && string.Equals(chart.LegendPosition, "t", StringComparison.Ordinal))
            {
                top = area.Top + area.Height * 0.34f;
                bottomReserve = 0.13f;
            }

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
            if (chart.Kind == ChartKind.Column && !chart.IsBar3D && chart.TextSizePt.HasValue && !chart.HasDataTable)
            {
                var textSize = (float)chart.TextSizePt.Value;
                left = area.Left + axisLabelMaxWidth + textSize * 1.5f;
                top = area.Top + (chart.Title == null ? textSize : (float)chart.TitleFont.Size + textSize * 2.5f);
                right = area.Right - textSize * 1.5f;
                bottom = area.Bottom - textSize * (chart.LegendPosition == "b" ? 4f : 2.5f);
            }
            if (UsesSideLegend(chart))
            {
                var legendWidth = SideLegendWidth(chart);
                if (chart.LegendPosition == "l") left += legendWidth;
                else right -= legendWidth;
                bottom = area.Bottom - (float)(chart.TextSizePt ?? 9d) * 2.5f;
            }
            if (!chart.HasDateCategoryAxis && Math.Abs(ResolveStandardCategoryAxisRotation(chart, right - left)) < 0.1f)
            {
                bottom -= ChartCategoryLabels.ExtraHeight(chart, _context.GetFontContext(ChartFont(LabelSizePt, false)));
            }
            return LineChartAutomaticLayout.Plot(area, chart, axisLabelMaxWidth,
                chart.HasDateCategoryAxis ? dateAxisLayout.RotationDeg : ResolveStandardCategoryAxisRotation(chart, right - left),
                new SKRect(left, top, right, bottom));
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
            var titleText = ResolveTitleText(chart);
            DrawTitleText(canvas, page, clip, titleText, clip.MidX, clip.MidY, titleFont,
                chart.TitleColor, TextAlign.Center, VAlign.Middle, chart.TitleTextOutlineColor);
        }

        private SKRect ResolveTitleBounds(SKRect area, ParsedChart chart, float centerY)
        {
            var titleFont = ResolveTitleFont(chart);
            var textWidth = MeasureText(ResolveTitleText(chart), titleFont);
            var horizontalPadding = Math.Max(3.5f, area.Width * 0.007f);
            var verticalPadding = 0.25f;
            if (chart.Kind == ChartKind.Line && !chart.IsLine3D && chart.HasTitleBorder && !chart.HasManualTitle)
            {
                horizontalPadding += (float)titleFont.Size * 0.05f;
                verticalPadding = 1.5f;
            }
            var width = Math.Min(area.Width * (chart.HasManualTitle ? 0.34f : 0.96f), textWidth + horizontalPadding * 2f);
            var left = chart.HasManualTitle
                ? area.Left + (float)(chart.TitleX * area.Width)
                : area.MidX - width * 0.5f;
            var right = area.MidX + width * 0.5f;
            if (chart.HasManualTitle)
            {
                right = left + width;
            }

            using (var paint = TextPaint(titleFont, chart.TitleColor))
            {
                var metrics = paint.FontMetrics;
                if (chart.HasManualTitle)
                {
                    var manualTop = area.Top + (float)(chart.TitleY * area.Height);
                    return new SKRect(left, manualTop, right, manualTop + metrics.Descent - metrics.Ascent + verticalPadding * 2f);
                }

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

            using (var grid = new SKPaint { Style = SKPaintStyle.Stroke, Color = chart.ValueGridlineColor, StrokeWidth = (float)Math.Max(0.5d, chart.ValueGridlineWidthPt), IsAntialias = false })
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

                    var label = FormatValue(value / chart.ValueDisplayUnitDivisor, chart.ValueFormatCode);
                    var width = MeasureText(label, LabelSizePt);
                    var clip = new SKRect(plot.Left - 8f - width, y - LabelSizePt, plot.Left - 1f, y + LabelSizePt);
                    if (!chart.ScatterValueAxisDeleted)
                    {
                        var labelX = chart.Kind == ChartKind.Scatter && chart.LegendPosition == null
                            ? plot.Left - 7.5f : plot.Left - 5f;
                        var labelColor = chart.AxisLabelColor;
                        if (value < 0d && !string.IsNullOrEmpty(chart.ValueFormatCode)
                            && chart.ValueFormatCode.IndexOf("[Red]", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            labelColor = SKColors.Red;
                        }
                        DrawText(canvas, page, clip, label, labelX, y, LabelSizePt, false,
                            labelColor, TextAlign.Right, VAlign.Middle, chart.AxisTextOutlineColor);
                    }
                }
            }

            ChartValueAxisLineRenderer.Draw(canvas, plot, chart, min, max, major);
        }

        private string ResolveTitleText(ParsedChart chart)
        {
            if (chart == null || string.IsNullOrEmpty(chart.Title) || !chart.TitleAllCaps)
            {
                return chart != null ? chart.Title : string.Empty;
            }

            var culture = chart.Culture ?? _chartCulture ?? CultureInfo.CurrentCulture;
            return chart.Title.ToUpper(culture);
        }

        private void DrawCategoryAxis(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            if (chart.Kind == ChartKind.Line && !chart.IsLine3D)
            {
                LineChartCategoryAxisRenderer.Draw(canvas, plot, chart, _dateSystem);
            }
            else
            {
                using (var axis = new SKPaint { Style = SKPaintStyle.Stroke, Color = GridColor, StrokeWidth = 1f, IsAntialias = false })
                {
                    canvas.DrawLine(plot.Left, plot.Bottom, plot.Right, plot.Bottom, axis);
                }
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
                    var lines = ChartCategoryLabels.Lines(label);
                    var advance = ChartCategoryLabels.LineAdvance(_context.GetFontContext(ChartFont(LabelSizePt, false)));
                    for (var line = 0; line < lines.Length; line++)
                    {
                        var width = MeasureText(lines[line], LabelSizePt);
                        var lineTop = plot.Bottom + line * advance;
                        var clip = new SKRect(x - width * 0.6f, lineTop + 2f, x + width * 0.6f,
                            Math.Min(visibleSlice.Bottom, lineTop + LabelSizePt * 2.2f));
                        DrawText(canvas, page, clip, lines[line], x, lineTop + 4f, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Center, VAlign.Top);
                    }
                }
            }
        }

        private void DrawScatterCategoryAxis(SKCanvas canvas, PageLayout page, SKRect plot, ParsedChart chart,
            double minimumY, double maximumY)
        {
            double minimum, maximum, major;
            ScatterChartAxisMath.ResolveXRange(chart, out minimum, out maximum, out major);
            if (major <= 0d)
            {
                return;
            }

            var axisY = chart.IsBubble && minimumY < 0d && maximumY > 0d
                ? ValueToY(0d, minimumY, maximumY, plot) : plot.Bottom;

            using (var axis = new SKPaint { Style = SKPaintStyle.Stroke,
                Color = chart.CategoryAxisLineColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.CategoryAxisLineWidthPt), IsAntialias = false })
            {
                if (chart.HasCategoryAxisLine)
                {
                    canvas.DrawLine(plot.Left, axisY, plot.Right, axisY, axis);
                }

                var count = (int)Math.Floor((maximum - minimum) / major + 0.0001d);
                if (chart.HasCategoryGridlines)
                {
                    using (var grid = new SKPaint { Style = SKPaintStyle.Stroke,
                        Color = chart.CategoryGridlineColor,
                        StrokeWidth = (float)Math.Max(0.5d, chart.CategoryGridlineWidthPt), IsAntialias = false })
                    {
                        for (var index = 0; index <= count; index++)
                        {
                            var x = ScatterChartAxisMath.ToScreen(minimum + index * major, minimum, maximum, plot.Left, plot.Right);
                            canvas.DrawLine(x, plot.Top, x, plot.Bottom, grid);
                        }
                    }
                }
                for (var index = 0; index <= count; index++)
                {
                    var value = minimum + index * major;
                    var x = ScatterChartAxisMath.ToScreen(value, minimum, maximum, plot.Left, plot.Right);
                    if (string.Equals(chart.CategoryAxisMajorTickMark, "out", StringComparison.Ordinal))
                    {
                        canvas.DrawLine(x, axisY, x, axisY + 3f, axis);
                    }

                    var labelValue = value == 0d && ChartXmlParser.IsDateFormat(chart.ScatterXFormatCode)
                        ? 2d : value;
                    var label = FormatValue(labelValue, chart.ScatterXFormatCode);
                    var clip = new SKRect(x - 24f, axisY + 4f, x + 24f, axisY + 20f);
                    if (!chart.ScatterCategoryAxisDeleted)
                    {
                        DrawText(canvas, page, clip, label, x, axisY + 6f,
                            LabelSizePt, false, chart.AxisLabelColor, TextAlign.Center, VAlign.Top,
                            chart.AxisTextOutlineColor);
                    }
                }
            }
        }

        private void DrawScatterSeries(SKCanvas canvas, SKRect plot, ParsedChart chart, double minimumY, double maximumY)
        {
            double minimumX, maximumX, majorX;
            ScatterChartAxisMath.ResolveXRange(chart, out minimumX, out maximumX, out majorX);
            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
            {
                var series = chart.Series[seriesIndex];
                var points = new List<SKPoint>();
                var count = Math.Min(series.XValues.Count, series.Values.Count);
                for (var index = 0; index < count; index++)
                {
                    if (!series.XValues[index].HasValue || !series.Values[index].HasValue)
                    {
                        continue;
                    }

                    var xValue = series.XValues[index].Value;
                    var yValue = series.Values[index].Value;
                    if (xValue < minimumX || xValue > maximumX || yValue < minimumY || yValue > maximumY)
                    {
                        continue;
                    }

                    points.Add(new SKPoint(
                        ScatterChartAxisMath.ToScreen(xValue, minimumX, maximumX, plot.Left, plot.Right),
                        ValueToY(yValue, minimumY, maximumY, plot)));
                }

                if (series.HasVisibleLine && points.Count > 1)
                {
                    using (var path = new SKPath())
                    using (var line = new SKPaint { Style = SKPaintStyle.Stroke, Color = series.LineColor,
                        StrokeWidth = (float)series.LineWidthPt, IsAntialias = true })
                    {
                        path.MoveTo(points[0]);
                        if (series.SmoothLine)
                        {
                            ScatterSmoothPathBuilder.Append(path, points);
                        }
                        else
                        {
                            for (var index = 1; index < points.Count; index++)
                            {
                                path.LineTo(points[index]);
                            }
                        }
                        canvas.DrawPath(path, line);
                    }
                }

                DrawSeriesMarkers(canvas, series, points);
            }
        }

        private void DrawHiddenAxisScatter(SKCanvas canvas, PageLayout page, SKRect area, ParsedChart chart)
        {
            var plot = new SKRect(area.Left + area.Width * 0.023f, area.Top + area.Height * 0.04f,
                area.Right - area.Width * 0.102f, area.Bottom - area.Height * 0.04f);
            double minimumY, maximumY, majorY;
            ComputeValueAxis(chart, out minimumY, out maximumY, out majorY);
            double minimumX, maximumX, majorX;
            ScatterChartAxisMath.ResolveXRange(chart, out minimumX, out maximumX, out majorX);
            var gridColor = SKColors.Black;
            using (var grid = new SKPaint { Style = SKPaintStyle.Stroke, Color = gridColor,
                StrokeWidth = 0.7f, IsAntialias = false })
            {
                var xSteps = (int)Math.Round((maximumX - minimumX) / majorX);
                for (var i = 0; i <= xSteps; i++)
                {
                    var x = ScatterChartAxisMath.ToScreen(minimumX + i * majorX, minimumX, maximumX, plot.Left, plot.Right);
                    canvas.DrawLine(x, plot.Top, x, plot.Bottom, grid);
                }
                var ySteps = (int)Math.Round((maximumY - minimumY) / majorY);
                for (var i = 0; i <= ySteps; i++)
                {
                    var y = ValueToY(minimumY + i * majorY, minimumY, maximumY, plot);
                    canvas.DrawLine(plot.Left, y, plot.Right, y, grid);
                }
            }

            DrawText(canvas, page, area, chart.Title, area.MidX, plot.Top + 3f,
                10f, true, SKColors.Black, TextAlign.Center, VAlign.Middle);
            DrawText(canvas, page, area, chart.CategoryAxisTitle, plot.MidX, plot.Bottom - 6f,
                10f, true, SKColors.Black, TextAlign.Center, VAlign.Middle);
            DrawRotatedText(canvas, page, area, chart.ValueAxisTitle, plot.Left + 7.5f, plot.MidY,
                10f, true, SKColors.Black, TextAlign.Center, VAlign.Middle, -90f);

            if (chart.Series.Count == 0)
            {
                return;
            }

            var series = chart.Series[0];
            var colors = new SKColor[] {
                new SKColor(0x68, 0x84, 0x40), new SKColor(0x7e, 0x9a, 0x49),
                new SKColor(0x90, 0xac, 0x52), new SKColor(0xa2, 0xbd, 0x65),
                new SKColor(0xb5, 0xcd, 0x89), new SKColor(0xc7, 0xda, 0xad),
                new SKColor(0xd5, 0xe3, 0xc6) };
            var count = Math.Min(series.XValues.Count, series.Values.Count);
            for (var i = 0; i < count; i++)
            {
                if (!series.XValues[i].HasValue || !series.Values[i].HasValue)
                {
                    continue;
                }
                var color = colors[Math.Min(i, colors.Length - 1)];
                var x = ScatterChartAxisMath.ToScreen(series.XValues[i].Value, minimumX, maximumX, plot.Left, plot.Right);
                var y = ValueToY(series.Values[i].Value, minimumY, maximumY, plot);
                using (var shadow = new SKPaint { Style = SKPaintStyle.Fill,
                    Color = new SKColor(0, 0, 0, 45), IsAntialias = true,
                    MaskFilter = SKMaskFilter.CreateBlur(SKBlurStyle.Normal, 1.4f) })
                using (var marker = new SKPaint { Style = SKPaintStyle.Fill, Color = color, IsAntialias = true })
                {
                    canvas.DrawCircle(x + 0.8f, y + 1f, 4.7f, shadow);
                    canvas.DrawCircle(x, y, 4.5f, marker);
                    var legendY = plot.Top + plot.Height * (0.294f + i * 0.0695f);
                    var legendX = plot.Right + area.Width * 0.039f;
                    canvas.DrawCircle(legendX, legendY, 3.5f, marker);
                    DrawText(canvas, page, area, FormatValue(series.XValues[i].Value, chart.ScatterXFormatCode),
                        legendX + 5f, legendY, 10f, false, SKColors.Black, TextAlign.Left, VAlign.Middle);
                }
            }
        }

        private void DrawAxisTitles(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            const float axisTitleSizePt = 10f;
            if (!string.IsNullOrEmpty(chart.CategoryAxisTitle))
            {
                DrawText(
                    canvas,
                    page,
                    area,
                    chart.CategoryAxisTitle,
                    plot.MidX,
                    plot.Bottom + (chart.Kind == ChartKind.Scatter ? 26f : 25f),
                    axisTitleSizePt,
                    false,
                    chart.AxisLabelColor,
                    TextAlign.Center,
                    VAlign.Middle);
            }

            if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
            {
                DrawRotatedText(
                    canvas,
                    page,
                    area,
                    chart.ValueAxisTitle,
                    plot.Left - 54f,
                    plot.MidY,
                    axisTitleSizePt,
                    false,
                    chart.AxisLabelColor,
                    TextAlign.Center,
                    VAlign.Middle,
                    -90f);
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

        private void DrawSeries(SKCanvas canvas, PageLayout page, SKRect plot, ParsedChart chart, double min, double max)
        {
            var pointCount = MaxPointCount(chart);

            if (chart.Kind == ChartKind.Radar)
            {
                return;
            }

            if (chart.Kind == ChartKind.Column)
            {
                DrawColumns(canvas, page, plot, chart, min, max, pointCount);
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
                        ChartSeriesStrokeRenderer.Draw(canvas, plot, chart, path, line);
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

        private void DrawColumns(SKCanvas canvas, PageLayout page, SKRect plot, ParsedChart chart, double min, double max, int pointCount)
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
                DrawStackedColumns(canvas, page, plot, chart, min, max, pointCount, slot, groupWidth, baselineY);
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
                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = ChartExplicitFormatting.PointColor(series, i), IsAntialias = true })
                    {
                        canvas.DrawRect(new SKRect(left, top, left + barWidth, bottom), fill);
                    }

                    ChartBarBorder.Draw(canvas, new SKRect(left, top, left + barWidth, bottom), series);
                    DrawColumnDataLabel(canvas, page, plot, chart, series, series.Values[i].Value,
                        left + barWidth * 0.5f, (top + bottom) * 0.5f, y);
                    DrawColumnErrorBars(canvas, plot, chart, series, s,
                        left + barWidth * 0.5f, series.Values[i].Value, min, max);
                }
            }
        }

        private void DrawColumnErrorBars(SKCanvas canvas, SKRect plot, ParsedChart chart, ChartSeries series,
            int seriesIndex, float centerX, double endValue, double min, double max)
        {
            if (series == null || !series.HasErrorBars
                || !string.Equals(series.ErrorValueType, "stdErr", StringComparison.Ordinal))
            {
                return;
            }

            var standardError = ResolveSeriesStandardError(chart, seriesIndex);
            if (standardError <= 0d)
            {
                return;
            }

            var top = ValueToY(endValue + standardError, min, max, plot);
            var bottom = ValueToY(endValue - standardError, min, max, plot);
            var capHalfWidth = 1.75f;
            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = series.ErrorBarColor,
                StrokeWidth = (float)Math.Max(0.5d, series.ErrorBarWidthPt),
                IsAntialias = false
            })
            {
                canvas.Save();
                canvas.ClipRect(plot);
                canvas.DrawLine(centerX, top, centerX, bottom, paint);
                if (series.ErrorBarsHaveEndCaps)
                {
                    canvas.DrawLine(centerX - capHalfWidth, top, centerX + capHalfWidth, top, paint);
                    canvas.DrawLine(centerX - capHalfWidth, bottom, centerX + capHalfWidth, bottom, paint);
                }
                canvas.Restore();
            }
        }

        private static double ResolveSeriesStandardError(ParsedChart chart, int seriesIndex)
        {
            if (chart == null || seriesIndex < 0 || seriesIndex >= chart.Series.Count)
            {
                return 0d;
            }

            var values = new List<double>();
            var series = chart.Series[seriesIndex];
            for (var pointIndex = 0; pointIndex < series.Values.Count; pointIndex++)
            {
                if (!series.Values[pointIndex].HasValue)
                {
                    continue;
                }

                var value = series.Values[pointIndex].Value;
                if (chart.IsPercentStacked)
                {
                    double total = 0d;
                    for (var otherIndex = 0; otherIndex < chart.Series.Count; otherIndex++)
                    {
                        var other = chart.Series[otherIndex];
                        if (pointIndex < other.Values.Count && other.Values[pointIndex].HasValue)
                        {
                            total += Math.Abs(other.Values[pointIndex].Value);
                        }
                    }

                    if (total <= 0d)
                    {
                        continue;
                    }

                    value /= total;
                }

                values.Add(value);
            }

            if (values.Count < 2)
            {
                return 0d;
            }

            double mean = 0d;
            for (var index = 0; index < values.Count; index++)
            {
                mean += values[index];
            }
            mean /= values.Count;

            double squaredDifferences = 0d;
            for (var index = 0; index < values.Count; index++)
            {
                var difference = values[index] - mean;
                squaredDifferences += difference * difference;
            }

            var sampleDeviation = Math.Sqrt(squaredDifferences / (values.Count - 1));
            return sampleDeviation / Math.Sqrt(values.Count);
        }

        private void DrawDataTable(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            if (chart == null || !chart.HasDataTable || chart.Series.Count == 0)
            {
                return;
            }

            var categoryCount = chart.Categories.Count;
            if (categoryCount == 0)
            {
                categoryCount = MaxPointCount(chart);
            }
            if (categoryCount == 0)
            {
                return;
            }

            var labelWidth = Math.Min(area.Width * 0.17f, 49f);
            var tableLeft = plot.Left - labelWidth;
            var tableTop = plot.Bottom + 24.5f;
            var rowHeight = 13.4f;
            var headerBottom = tableTop + rowHeight;
            var tableBottom = headerBottom + chart.Series.Count * rowHeight;
            var cellWidth = plot.Width / categoryCount;

            using (var border = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.DataTableLineColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.DataTableLineWidthPt),
                IsAntialias = false
            })
            {
                if (chart.DataTableShowHorizontalBorders)
                {
                    canvas.DrawLine(plot.Left, tableTop, plot.Right, tableTop, border);
                    for (var row = 0; row <= chart.Series.Count; row++)
                    {
                        var y = headerBottom + row * rowHeight;
                        canvas.DrawLine(tableLeft, y, plot.Right, y, border);
                    }
                }

                if (chart.DataTableShowVerticalBorders)
                {
                    canvas.DrawLine(tableLeft, headerBottom, tableLeft, tableBottom, border);
                    for (var column = 0; column <= categoryCount; column++)
                    {
                        var x = plot.Left + column * cellWidth;
                        canvas.DrawLine(x, tableTop, x, tableBottom, border);
                    }
                }

                if (chart.DataTableShowOutline)
                {
                    canvas.DrawLine(plot.Left, tableTop, plot.Right, tableTop, border);
                    canvas.DrawLine(tableLeft, tableBottom, plot.Right, tableBottom, border);
                    canvas.DrawLine(plot.Right, tableTop, plot.Right, tableBottom, border);
                }
            }

            for (var column = 0; column < categoryCount; column++)
            {
                var category = ResolveCategoryLabel(chart, column);
                var centerX = plot.Left + (column + 0.5f) * cellWidth;
                DrawText(canvas, page, new SKRect(plot.Left + column * cellWidth, tableTop,
                    plot.Left + (column + 1) * cellWidth, headerBottom), category, centerX,
                    tableTop + rowHeight * 0.5f, LabelSizePt, false, chart.DataTableTextColor,
                    TextAlign.Center, VAlign.Middle);
            }

            for (var row = 0; row < chart.Series.Count; row++)
            {
                var seriesIndex = chart.Series.Count - 1 - row;
                var series = chart.Series[seriesIndex];
                var centerY = headerBottom + (row + 0.5f) * rowHeight;
                var textLeft = tableLeft + 3f;
                if (chart.DataTableShowKeys)
                {
                    using (var key = new SKPaint { Style = SKPaintStyle.Fill, Color = series.Color, IsAntialias = false })
                    {
                        canvas.DrawRect(new SKRect(tableLeft + 3f, centerY - 2.3f,
                            tableLeft + 7.6f, centerY + 2.3f), key);
                    }
                    textLeft = tableLeft + 10f;
                }

                DrawText(canvas, page, new SKRect(tableLeft, centerY - rowHeight * 0.5f,
                    plot.Left, centerY + rowHeight * 0.5f), series.Name, textLeft, centerY,
                    LabelSizePt, false, chart.DataTableTextColor, TextAlign.Left, VAlign.Middle);

                for (var column = 0; column < categoryCount; column++)
                {
                    if (column >= series.Values.Count || !series.Values[column].HasValue)
                    {
                        continue;
                    }

                    var formatCode = !string.IsNullOrEmpty(series.ValueFormatCode)
                        ? series.ValueFormatCode
                        : chart.ValueFormatCode;
                    var value = FormatValue(series.Values[column].Value, formatCode);
                    var centerX = plot.Left + (column + 0.5f) * cellWidth;
                    DrawText(canvas, page, new SKRect(plot.Left + column * cellWidth,
                        centerY - rowHeight * 0.5f, plot.Left + (column + 1) * cellWidth,
                        centerY + rowHeight * 0.5f), value, centerX, centerY, LabelSizePt,
                        false, chart.DataTableTextColor, TextAlign.Center, VAlign.Middle);
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
                    ChartBarBorder.Draw(canvas, new SKRect(left, top, right, top + barHeight), series);
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

                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = series.Color, IsAntialias = true })
                    {
                        canvas.DrawRect(new SKRect(left, top, right, top + barHeight), fill);
                    }
                    ChartBarBorder.Draw(canvas, new SKRect(left, top, right, top + barHeight), series);
                }
            }
        }

        private void DrawStackedColumns(SKCanvas canvas, PageLayout page, SKRect plot, ParsedChart chart, double min, double max, int pointCount, float slot, float barWidth, float baselineY)
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

                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = ChartExplicitFormatting.PointColor(series, i), IsAntialias = true })
                    {
                        canvas.DrawRect(new SKRect(left, top, left + barWidth, bottom), fill);
                    }

                    ChartBarBorder.Draw(canvas, new SKRect(left, top, left + barWidth, bottom), series);
                    DrawColumnDataLabel(canvas, page, plot, chart, series, series.Values[i].Value,
                        left + barWidth * 0.5f, (top + bottom) * 0.5f, endY);
                    DrawColumnErrorBars(canvas, plot, chart, series, s,
                        left + barWidth * 0.5f, endValue, min, max);
                }
            }
        }

        private void DrawColumnDataLabel(SKCanvas canvas, PageLayout page, SKRect plot, ParsedChart chart,
            ChartSeries series, double value, float centerX, float centerY, float endY)
        {
            if (series == null || !series.ShowValueDataLabels)
            {
                return;
            }

            if (!string.IsNullOrEmpty(series.DataLabelPosition)
                && !string.Equals(series.DataLabelPosition, "ctr", StringComparison.Ordinal)
                && !string.Equals(series.DataLabelPosition, "inEnd", StringComparison.Ordinal))
            {
                return;
            }

            var formatCode = !string.IsNullOrEmpty(series.ValueFormatCode)
                ? series.ValueFormatCode
                : chart.ValueFormatCode;
            var label = FormatValue(value, formatCode);
            var size = (float)(series.DataLabelSizePt ?? LabelSizePt);
            if (series.DataLabelPosition == "inEnd") centerY = endY + (value < 0d ? -1f : 1f) * (size * 0.65f + 2f);
            DrawText(canvas, page, plot, label, centerX, centerY, size, false,
                series.DataLabelColor, TextAlign.Center, VAlign.Middle);
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

            var sideLegend = UsesSideLegend(chart);
            var y = LegendCenterY(area, plot, chart);
            var swatch = ResolveLegendSwatchWidth(chart);
            var gap = ResolveLegendTextGap(chart);
            var itemGap = ResolveLegendItemGap(chart);

            var widths = new float[legendItemCount];
            var total = 0f;
            for (var i = 0; i < legendItemCount; i++)
            {
                widths[i] = swatch + gap + MeasureText(LegendItemLabel(chart, i), LabelSizePt);
                total += widths[i] + (i > 0 ? itemGap : 0f);
            }

            var x = ResolveLegendLeft(area, chart, total);
            var legendTextBottom = y;
            var legendSwatchHeight = chart.HasManualLegend ? 4.5f : IsCompactColumnLegend(chart) ? 5f : 5.5f;
            var legendSwatchVerticalOffset = chart.HasManualLegend ? -1f : IsCompactColumnLegend(chart) ? -0.5f : 0f;
            using (var metricsPaint = TextPaint(ChartFont(LabelSizePt, false), chart.LegendTextColor))
            {
                var metrics = metricsPaint.FontMetrics;
                legendTextBottom = y + (metrics.Descent - metrics.Ascent) / 2f - 1.4f;
            }

            var lineHeight = LabelSizePt * 1.7f;
            if (sideLegend)
            {
                var textBottomOffset = legendTextBottom - y;
                x = chart.LegendPosition == "l" ? area.Left + LabelSizePt : area.Right - SideLegendWidth(chart);
                y = area.MidY - lineHeight * (legendItemCount - 1) * 0.5f;
                legendTextBottom = y + textBottomOffset;
            }

            for (var i = 0; i < legendItemCount; i++)
            {
                var label = LegendItemLabel(chart, i);
                var color = LegendItemColor(chart, i);
                if (chart.Kind == ChartKind.Column || chart.Kind == ChartKind.Bar || chart.Kind == ChartKind.Pie
                    || chart.Kind == ChartKind.Area || (chart.Kind == ChartKind.Line && chart.IsLine3D))
                {
                    var rectBottom = legendTextBottom + legendSwatchVerticalOffset;
                    var rectTop = rectBottom - legendSwatchHeight;
                    var swatchRect = new SKRect(x, rectTop, x + legendSwatchHeight, rectBottom);
                    var patternDrawn = false;
                    if (chart.Kind == ChartKind.Area)
                    {
                        var seriesIndex = ResolveLegendSeriesIndex(chart, i);
                        if (seriesIndex >= 0 && seriesIndex < chart.Series.Count)
                        {
                            patternDrawn = ChartPatternRenderer.DrawRect(canvas, swatchRect, chart.Series[seriesIndex], color);
                        }
                    }

                    if (!patternDrawn)
                    {
                        using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = color, IsAntialias = true })
                        {
                            canvas.DrawRect(swatchRect, fill);
                        }
                    }
                }
                else if (chart.IsBubble)
                {
                    var legendSeriesIndex = ResolveLegendSeriesIndex(chart, i);
                    if (legendSeriesIndex >= 0 && legendSeriesIndex < chart.Series.Count
                        && chart.Series[legendSeriesIndex].Bubble3D)
                    {
                        BubbleSphereRenderer.Draw(canvas, new SKPoint(x + swatch * 0.5f, y), 2.45f, color);
                    }
                    else
                    {
                        using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = color, IsAntialias = true })
                        {
                            canvas.DrawCircle(x + swatch * 0.5f, y, 2.45f, fill);
                        }
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
                    var legendSeriesIndex = ResolveLegendSeriesIndex(chart, i);
                    var showLine = chart.Kind != ChartKind.Scatter
                        || legendSeriesIndex >= 0 && legendSeriesIndex < chart.Series.Count && chart.Series[legendSeriesIndex].HasVisibleLine;
                    if (showLine)
                    {
                        var legendLineWidth = LineChartDecorationLayout.UsesAutomaticBottomLegend(chart) ? 2.25f : 2.5f;
                        if (chart.Kind == ChartKind.Scatter && legendSeriesIndex >= 0 && legendSeriesIndex < chart.Series.Count)
                        {
                            legendLineWidth = (float)chart.Series[legendSeriesIndex].LineWidthPt;
                        }
                        using (var line = new SKPaint { Style = SKPaintStyle.Stroke, Color = color, StrokeWidth = legendLineWidth, IsAntialias = true, StrokeCap = SKStrokeCap.Round })
                        {
                            canvas.DrawLine(x, y, x + swatch, y, line);
                        }
                    }

                    if (chart.Kind == ChartKind.Line || chart.Kind == ChartKind.Scatter)
                    {
                        if (legendSeriesIndex >= 0 && legendSeriesIndex < chart.Series.Count && chart.Series[legendSeriesIndex].MarkerVisible)
                        {
                            var series = chart.Series[legendSeriesIndex];
                            using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = series.MarkerFillColor, IsAntialias = true })
                            using (var stroke = new SKPaint { Style = SKPaintStyle.Stroke, Color = series.MarkerStrokeColor,
                                StrokeWidth = (float)Math.Max(0.75d, series.MarkerStrokeWidthPt), IsAntialias = true })
                            {
                                DrawMarker(canvas, new SKPoint(x + swatch * 0.5f, y),
                                    Math.Max(3f, series.MarkerSize) * 0.5f, series.MarkerSymbol, fill, stroke);
                            }
                        }
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
                DrawText(canvas, page, clip, label, x + swatch + gap, y, LabelSizePt, false,
                    chart.LegendTextColor, TextAlign.Left, VAlign.Middle, chart.LegendTextOutlineColor);
                if (sideLegend)
                {
                    y += lineHeight;
                    legendTextBottom += lineHeight;
                }
                else x += widths[i] + itemGap;
            }
        }

        private SKRect GetLegendBounds(SKRect area, SKRect plot, ParsedChart chart)
        {
            if (UsesSideLegend(chart))
            {
                var width = SideLegendWidth(chart);
                var left = chart.LegendPosition == "l" ? area.Left + LabelSizePt : area.Right - width;
                var height = LegendItemCount(chart) * LabelSizePt * 1.7f;
                return new SKRect(left, area.MidY - height * 0.5f, left + width, area.MidY + height * 0.5f);
            }
            var y = LegendCenterY(area, plot, chart);
            var swatch = ResolveLegendSwatchWidth(chart);
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

            var x = ResolveLegendLeft(area, chart, total);
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

            if (chart.Kind == ChartKind.Column && (chart.IsStacked || chart.IsPercentStacked) && UsesSideLegend(chart))
            {
                return chart.Series.Count - 1 - legendIndex;
            }

            return legendIndex;
        }

        private static bool UsesSideLegend(ParsedChart chart)
        {
            return !chart.HasManualLegend && !chart.IsBar3D
                && (chart.Kind == ChartKind.Column || chart.Kind == ChartKind.Bar)
                && (chart.LegendPosition == "r" || chart.LegendPosition == "l" || chart.LegendPosition == "tr");
        }

        private float SideLegendWidth(ParsedChart chart)
        {
            var textWidth = 0f;
            var size = (float)(chart.TextSizePt ?? 9d);
            for (var i = 0; i < LegendItemCount(chart); i++)
                textWidth = Math.Max(textWidth, MeasureText(LegendItemLabel(chart, i), size));
            return textWidth + ResolveLegendSwatchWidth(chart) + ResolveLegendTextGap(chart) + size;
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
            if (chart != null && chart.HasManualLegend)
            {
                return area.Top + (float)(chart.LegendY + chart.LegendH * 0.5d) * area.Height - 1.67f;
            }

            if (chart != null && string.Equals(chart.LegendPosition, "t", StringComparison.Ordinal))
            {
                return area.Top + area.Height * 0.23f;
            }

            if (chart != null && chart.Kind == ChartKind.Pie && !chart.IsOfPie)
            {
                return area.Bottom - area.Height * 0.085f;
            }

            var position = 0.68f;
            var dateAxisLayout = ResolveDateAxisLayout(chart, plot.Width);
            if (chart.HasDateCategoryAxis && Math.Abs(dateAxisLayout.RotationDeg) > 0.1f)
            {
                position = 0.86f;
            }
            if (chart.IsLine3D)
            {
                position = 0.76f;
            }

            var centerY = plot.Bottom + (area.Bottom - plot.Bottom) * position;
            if (chart != null && chart.IsBubble)
            {
                centerY -= 6f;
            }
            if (chart != null && chart.Kind == ChartKind.Scatter
                && !string.IsNullOrEmpty(chart.CategoryAxisTitle))
            {
                centerY = area.Bottom - 15.7f;
            }
            if (chart != null && chart.HasDataTable)
            {
                centerY = area.Top + area.Height * 0.919f;
            }
            else if (chart != null && chart.Kind == ChartKind.Column && chart.IsBar3D)
            {
                centerY = area.Top + area.Height * 0.93f;
            }
            if (IsCompactColumnLegend(chart))
            {
                centerY -= 1f;
            }
            if (chart != null && chart.IsOfPie && string.Equals(chart.OfPieType, "bar", StringComparison.OrdinalIgnoreCase))
            {
                centerY -= 5f;
            }

            return LineChartDecorationLayout.LegendCenterY(area, chart, centerY);
        }

        private float ResolveLegendTextGap(ParsedChart chart)
        {
            if (UsesLineScatterLegend(chart))
            {
                return 2f;
            }

            if (IsCompactColumnLegend(chart))
            {
                return 2f;
            }

            if (chart != null && chart.Kind == ChartKind.Pie)
            {
                return 0f;
            }

            if (chart != null && chart.Kind == ChartKind.Area)
            {
                return 0f;
            }

            return LineChartDecorationLayout.TextGap(chart, 1f);
        }

        private float ResolveLegendItemGap(ParsedChart chart)
        {
            if (chart != null && chart.Kind == ChartKind.Scatter)
            {
                if (UsesLineScatterLegend(chart))
                {
                    return 9.5f;
                }

                return chart.HasManualLegend ? 6.9f : 4.6f;
            }

            if (chart != null && chart.IsLine3D)
            {
                return 4.6f;
            }

            if (IsCompactColumnLegend(chart))
            {
                return 7.45f;
            }

            if (chart != null && chart.Kind == ChartKind.Pie)
            {
                return chart.IsDoughnut ? 7.1f : 3.6f;
            }

            if (chart != null && chart.Kind == ChartKind.Area)
            {
                return 6f;
            }

            return LineChartDecorationLayout.ItemGap(chart, 12.6f);
        }

        private float ResolveLegendSwatchWidth(ParsedChart chart)
        {
            if (UsesLineScatterLegend(chart))
            {
                return 19.2f;
            }

            if (chart != null && chart.IsDoughnut)
            {
                return 7f;
            }

            if (chart != null && chart.HasManualLegend)
            {
                return 8.25f;
            }

            if (IsCompactColumnLegend(chart))
            {
                return 5f;
            }

            return LineChartDecorationLayout.SwatchWidth(chart, 9f);
        }

        private static bool UsesLineScatterLegend(ParsedChart chart)
        {
            if (chart == null || chart.Kind != ChartKind.Scatter || chart.Series == null)
            {
                return false;
            }

            for (var i = 0; i < chart.Series.Count; i++)
            {
                if (chart.Series[i].HasVisibleLine)
                {
                    return true;
                }
            }

            return false;
        }

        private float ResolveLegendHorizontalOffset(ParsedChart chart)
        {
            if (chart != null && chart.Kind == ChartKind.Scatter)
            {
                return 1.5f;
            }

            if (chart != null && chart.IsDoughnut)
            {
                return 2.5f;
            }

            if (chart != null && chart.IsLine3D)
            {
                return 3.5f;
            }

            if (IsCompactColumnLegend(chart))
            {
                return 2.25f;
            }

            return LineChartDecorationLayout.HorizontalOffset(chart, 0f);
        }

        private float ResolveLegendLeft(SKRect area, ParsedChart chart, float totalWidth)
        {
            if (chart != null && chart.HasManualLegend)
            {
                var center = area.Left + (float)(chart.LegendX + chart.LegendW * 0.5d) * area.Width;
                return center - totalWidth * 0.5f + (chart.Kind == ChartKind.Scatter ? 2.3f : 5.75f);
            }

            return area.MidX - totalWidth * 0.5f + ResolveLegendHorizontalOffset(chart);
        }

        private bool IsCompactColumnLegend(ParsedChart chart)
        {
            return chart != null
                && chart.Kind == ChartKind.Column
                && !chart.IsPercentStacked
                && !chart.IsBar3D
                && !chart.HasManualLegend
                && chart.Title != null
                && chart.LegendPosition != null;
        }

        private void DrawPie(SKCanvas canvas, PageLayout page, SKRect area, ParsedChart chart)
        {
            var titlePosition = 0.08f - (chart.IsOfPie ? ResolveOfPieWideLayoutFactor(area) * 0.025f : 0f);
            DrawTitle(canvas, page, area, chart, area.Top + area.Height * titlePosition);
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

            if (chart.IsDoughnut)
            {
                var doughnutRect = DoughnutChartRenderer.Draw(canvas, area, chart);
                DrawPlotAreaBorder(canvas, page, area, ResolvePieBorderRect(area, chart, doughnutRect), chart);
                DrawLegend(canvas, page, area, doughnutRect, chart);
                return;
            }

            if (chart.IsPie3D)
            {
                var pointColors = new List<SKColor>();
                for (var i = 0; i < values.Count; i++)
                {
                    pointColors.Add(PiePointColor(chart, i));
                }

                var ellipse = Pie3DChartRenderer.Draw(canvas, area, chart, values, pointColors, sum);
                DrawPlotAreaBorder(canvas, page, area, ResolvePieBorderRect(area, chart, ellipse), chart);
                DrawLegend(canvas, page, area, ellipse, chart);
                return;
            }

            var size = Math.Min(area.Width, area.Height) * 0.635f;
            var rect = new SKRect(area.MidX - size / 2f, area.Top + area.Height * 0.19f, area.MidX + size / 2f, area.Top + area.Height * 0.19f + size);
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
                    if (i < chart.PiePointLineWidthsPt.Count && chart.PiePointLineWidthsPt[i] > 0d)
                    {
                        using (var outline = new SKPaint { Style = SKPaintStyle.Stroke,
                            Color = chart.PiePointLineColors[i],
                            StrokeWidth = (float)chart.PiePointLineWidthsPt[i], IsAntialias = true })
                        {
                            canvas.DrawPath(path, outline);
                        }
                    }
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
            var wideLayout = ResolveOfPieWideLayoutFactor(area);
            var baseMainSize = Math.Min(area.Width, area.Height) * (0.515f + largerChartBonus * 0.02f);
            var maxMainSize = Math.Min(plotRect.Width * (0.64f + largerChartBonus * 0.03f), plotRect.Height * 0.96f);
            var mainSize = Math.Min(baseMainSize, maxMainSize) * (1.22f - 0.12f * wideLayout);
            var secondSize = mainSize * (float)Math.Max(0.35d, Math.Min(1.2d, chart.OfPieSecondPieSizePercent / 100d));
            var gapRatio = (float)Math.Max(0.1d, Math.Min(3d, chart.OfPieGapWidthPercent / 100d));
            var centerGap = 44.465f + 32.643f * gapRatio + 26.5f * wideLayout;
            var isBarOfPie = string.Equals(chart.OfPieType, "bar", StringComparison.OrdinalIgnoreCase);
            var wideBarScale = isBarOfPie ? 1f + 0.25f * wideLayout : 1f;
            mainSize *= wideBarScale;
            secondSize *= wideBarScale;
            centerGap *= wideBarScale;
            if (isBarOfPie)
            {
                centerGap *= 0.5f;
            }

            var secondWidth = isBarOfPie ? secondSize * 0.5f : secondSize;
            var groupWidth = mainSize + centerGap + secondWidth;
            var groupWidthLimit = plotRect.Width * (0.94f + largerChartBonus * 0.025f + wideLayout * 0.055f);
            if (groupWidth > groupWidthLimit)
            {
                var scale = groupWidthLimit / groupWidth;
                mainSize *= scale;
                secondSize *= scale;
                secondWidth *= scale;
                centerGap *= scale;
                groupWidth = mainSize + centerGap + secondWidth;
            }

            var groupHeight = Math.Max(mainSize, secondSize);
            var left = plotRect.MidX - groupWidth / 2f - (isBarOfPie ? 2f * wideLayout : 0f);
            var top = plotRect.MidY + area.Height * 0.018f - groupHeight / 2f;
            var mainRect = new SKRect(left, top + (mainSize < secondSize ? (secondSize - mainSize) * 0.5f : 0f), left + mainSize, top + (mainSize < secondSize ? (secondSize - mainSize) * 0.5f : 0f) + mainSize);
            var secondLeft = mainRect.Right + centerGap;
            var secondRect = new SKRect(secondLeft, top + (secondSize < mainSize ? (mainSize - secondSize) * 0.5f : 0f), secondLeft + secondWidth, top + (secondSize < mainSize ? (mainSize - secondSize) * 0.5f : 0f) + secondSize);

            float aggregateStart;
            float aggregateSweep;
            DrawPieSlices(canvas, mainRect, primaryValues, primaryColors, 70f, out aggregateStart, out aggregateSweep);

            if (isBarOfPie)
            {
                var segmentTop = secondRect.Top;
                for (var i = 0; i < secondaryValues.Count; i++)
                {
                    var segmentBottom = i == secondaryValues.Count - 1
                        ? secondRect.Bottom
                        : segmentTop + secondRect.Height * (float)(secondaryValues[i] / secondarySum);
                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = secondaryColors[i], IsAntialias = true })
                    {
                        canvas.DrawRect(new SKRect(secondRect.Left, segmentTop, secondRect.Right, segmentBottom), fill);
                    }

                    segmentTop = segmentBottom;
                }
            }
            else
            {
                float secondaryStart;
                float secondarySweep;
                DrawPieSlices(canvas, secondRect, secondaryValues, secondaryColors, aggregateStart + 145f, out secondaryStart, out secondarySweep);
            }

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
            var wideLayout = ResolveOfPieWideLayoutFactor(area);
            var borderRect = ExpandRectAboutCenter(plotRect, 1f + 0.116f * largerChartBonus, 1.29f);

            var topLimit = area.Top + (string.IsNullOrEmpty(chart != null ? chart.Title : null) ? area.Height * 0.06f : area.Height * (0.12f + 0.013f * wideLayout));
            if (borderRect.Top < topLimit)
            {
                borderRect.Top = topLimit;
            }

            var bottomLimit = area.Bottom - (chart != null && chart.LegendPosition != null ? area.Height * (0.18f - 0.054f * wideLayout) : area.Height * 0.05f);
            if (borderRect.Bottom > bottomLimit)
            {
                borderRect.Bottom = bottomLimit;
            }

            var sideInset = area.Width * 0.026f * wideLayout;
            return InsetAndClamp(area, Math.Max(borderRect.Left, area.Left + sideInset), borderRect.Top,
                Math.Min(borderRect.Right, area.Right - sideInset), borderRect.Bottom);
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
            if (string.Equals(chart.OfPieType, "bar", StringComparison.OrdinalIgnoreCase))
            {
                topSecond = new SKPoint(secondRect.Left, secondRect.Top);
                bottomSecond = new SKPoint(secondRect.Left, secondRect.Bottom);
            }
            else if (!TryResolveOuterTangents(mainRect, secondRect, out topMain, out topSecond, out bottomMain, out bottomSecond))
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

            topMain = PieArcPoint(mainRect, aggregateStart);
            bottomMain = PieArcPoint(mainRect, aggregateStart + aggregateSweep);

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
            if (chart.Kind == ChartKind.Column && chart.CrossBetween)
            {
                DrawColumnDateCategoryAxis(canvas, page, area, plot, chart);
                return;
            }

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
            var boundaryGridlines = LineChartCategoryAxisRenderer.DrawBoundaryGridlines(canvas, plot, chart, _dateSystem);
            var visibleSlice = ResolveVisibleChartSlice(page, area);
            var current = firstTick;
            var rotatedClip = new SKRect(
                Math.Max(plot.Left - 36f, visibleSlice.Left),
                plot.Bottom + 2f,
                Math.Min(plot.Right + 18f, visibleSlice.Right),
                visibleSlice.Bottom);
            while (current <= lastTick + 0.1d)
            {
                var x = ChartDateAxisMath.DateCategoryX(current, plotMinDate, plotMaxDate, plot, chart, _dateSystem);
                if (chart.HasCategoryGridlines && !boundaryGridlines)
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

                    var automaticDiagonal = LineChartAutomaticLayout.Applies(chart) && Math.Abs(layout.RotationDeg + 45f) < 0.1f;
                    var labelX = automaticDiagonal ? x - LabelSizePt * 0.5f : x + 2f;
                    DrawRotatedText(canvas, page, rotatedClip, label, labelX, plot.Bottom + (automaticDiagonal ? 4f : 6f), LabelSizePt, false, chart.AxisLabelColor, TextAlign.Right, VAlign.Top, layout.RotationDeg);
                }
                else
                {
                    DrawText(canvas, page, area, label, x, plot.Bottom + 4f, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Center, VAlign.Top);
                }
                current = ChartDateAxisMath.AddMonthsSerial(current, layout.MonthStep, _dateSystem);
            }
        }

        private void DrawColumnDateCategoryAxis(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart)
        {
            var count = Math.Max(chart.Categories.Count, chart.CategoryValues.Count);
            if (count <= 0)
            {
                return;
            }

            if (chart.HasCategoryGridlines)
            {
                using (var grid = new SKPaint
                {
                    Style = SKPaintStyle.Stroke,
                    Color = chart.CategoryGridlineColor,
                    StrokeWidth = (float)Math.Max(0.5d, chart.CategoryGridlineWidthPt),
                    IsAntialias = false
                })
                {
                    for (var i = 0; i <= count; i++)
                    {
                        var x = ColumnDateCategoryBoundaryX(i, count, plot);
                        canvas.DrawLine(x, plot.Top, x, plot.Bottom, grid);
                    }
                }
            }

            if (_suppressCategoryAxisLabels)
            {
                return;
            }

            for (var i = 0; i < count; i++)
            {
                var label = ResolveCategoryLabel(chart, i);
                var x = CategoryX(i, count, plot, true);
                DrawText(canvas, page, area, label, x, plot.Bottom + 4f, LabelSizePt, false, chart.AxisLabelColor, TextAlign.Center, VAlign.Top);
            }
        }

        private static float ColumnDateCategoryBoundaryX(int index, int count, SKRect plot)
        {
            if (count <= 0)
            {
                return plot.Left;
            }

            return plot.Left + (float)index / count * plot.Width;
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
                    return ChartDateAxisMath.DateCategoryX(serial, plotMinDate, plotMaxDate, plot, chart, _dateSystem);
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

            if (chart.IsBubble)
            {
                BubbleChartMath.ResolveYRange(chart, dmin, dmax, out min, out max, out major);
                return;
            }

            if (chart != null && chart.IsPercentStacked)
            {
                min = chart.AxisMin ?? 0d;
                max = chart.AxisMax ?? 1d;
                var automaticMajor = chart.Kind == ChartKind.Column && chart.IsBar3D
                    && string.Equals(chart.BarShape, "cylinder", StringComparison.Ordinal)
                    ? 0.5d
                    : 0.2d;
                major = chart.MajorUnit ?? automaticMajor;
                if (major <= 0d)
                {
                    major = automaticMajor;
                }

                if (max <= min)
                {
                    max = min + major;
                }

                return;
            }

            min = chart.AxisMin ?? (dmin >= 0d ? 0d : -NiceNumber(-dmin));
            var compactArea3D = chart != null
                && chart.IsArea3D
                && chart.Series.Count > Math.Max(chart.Categories.Count, chart.CategoryValues.Count);
            var labeledColumn3D = false;
            if (chart != null && chart.Kind == ChartKind.Column && chart.IsBar3D)
            {
                for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
                {
                    if (chart.Series[seriesIndex].ShowValueDataLabels)
                    {
                        labeledColumn3D = true;
                        break;
                    }
                }
            }
            var automaticMajorIntervals = chart != null && chart.IsLine3D
                ? 4d
                : chart != null && chart.Kind == ChartKind.Scatter
                    && !string.IsNullOrEmpty(chart.ValueFormatCode) && chart.ValueFormatCode.Contains("%")
                    ? 6d
                : compactArea3D && chart.IsStacked
                ? (chart.PerspectiveDeg >= 100d ? 2d : 3d)
                : chart != null && chart.Kind == ChartKind.Column && chart.IsBar3D
                    ? (labeledColumn3D ? 7d : 4d)
                    : 8d;
            major = chart.MajorUnit ?? NiceNumber((dmax - min) / automaticMajorIntervals);
            if (major <= 0d) major = 1d;

            if (chart.AxisMax.HasValue)
            {
                max = chart.AxisMax.Value;
            }
            else if (chart.IsStacked && !chart.IsBar3D && min == 0d && dmax > 0d
                && (chart.Kind == ChartKind.Column || chart.Kind == ChartKind.Bar))
            {
                // Positive 2-D stacks reserve 5% above the total before rounding to a tick.
                // Excel probes: 571 -> 600, 572 -> 700 (major unit 100).
                max = Math.Ceiling(dmax * 1.05d / major) * major;
            }
            else
            {
                max = Math.Ceiling((dmax - min) / major) * major + min;
                // Radar charts use the rounded outer ring directly instead of adding
                // an extra headroom band above the data maximum.
                var denseBar3DGrid = chart != null && chart.IsBar3D && chart.HasMinorValueGridlines;
                var standardDateBar3D = chart != null && chart.IsBar3D
                    && chart.HasDateCategoryAxis && chart.PerspectiveDeg < 100d;
                if ((chart == null || chart.Kind != ChartKind.Radar) && !compactArea3D && !denseBar3DGrid
                    && !standardDateBar3D
                    && max - dmax < 0.15d * major) max += major;
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
                var label = FormatValue(value / chart.ValueDisplayUnitDivisor, chart.ValueFormatCode);
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
                if (chart.Kind == ChartKind.Scatter)
                {
                    canvas.DrawRect(plot, paint);
                    return;
                }

                if (chart.Kind == ChartKind.Pie)
                {
                    canvas.DrawLine(plot.Left, plot.Top, plot.Right, plot.Top, paint);
                    canvas.DrawLine(plot.Left, plot.Bottom, plot.Right, plot.Bottom, paint);
                    canvas.DrawLine(plot.Left, plot.Top, plot.Left, plot.Bottom, paint);
                    canvas.DrawLine(plot.Right, plot.Top, plot.Right, plot.Bottom, paint);
                    return;
                }

                var drawRightEdge = IsLastHorizontalChartSlice(page, area);
                var verticalStrokeWidth = chart.Kind == ChartKind.Line ? paint.StrokeWidth : Math.Max(0.5f, paint.StrokeWidth * 0.7f);
                canvas.DrawLine(plot.Left, plot.Top, plot.Right, plot.Top, paint);
                paint.StrokeWidth = verticalStrokeWidth;
                canvas.DrawLine(plot.Left, plot.Top, plot.Left, plot.Bottom, paint);
                if (drawRightEdge)
                {
                    canvas.DrawLine(plot.Right, plot.Top, plot.Right, plot.Bottom, paint);
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

        private static float ResolveOfPieWideLayoutFactor(SKRect area)
        {
            return Math.Max(0f, Math.Min(1f, (Math.Min(area.Width, area.Height) - 200f) / 70f));
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
            return ChartValueAxisIntervals.ResolveMajor(chart, plot, major);
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
                return LineChartAutomaticLayout.Applies(chart) && !chart.HasManualPlot
                    ? plot.Right : plot.Right + StandardDateAxisGridlineRightExtensionPt;
            }

            if (chart != null
                && chart.Kind == ChartKind.Line
                && chart.IsStacked
                && !chart.IsPercentStacked
                && chart.HasDateCategoryAxis
                && chart.LegendPosition != null)
            {
                return LineChartAutomaticLayout.Applies(chart) && !chart.HasManualPlot
                    ? plot.Right : plot.Right + StackedDateAxisGridlineRightExtensionPt;
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
            DrawText(canvas, page, clipRect, text, x, y, sizePt, bold, false, color, hAlign, vAlign, null);
        }

        private void DrawText(SKCanvas canvas, PageLayout page, SKRect clipRect, string text, float x, float y,
            float sizePt, bool bold, SKColor color, TextAlign hAlign, VAlign vAlign, SKColor? outlineColor)
        {
            DrawText(canvas, page, clipRect, text, x, y, sizePt, bold, false, color, hAlign, vAlign, outlineColor);
        }

        private void DrawText(SKCanvas canvas, PageLayout page, SKRect clipRect, string text, float x, float y, float sizePt, bool bold, bool italic, SKColor color, TextAlign hAlign, VAlign vAlign)
        {
            DrawText(canvas, page, clipRect, text, x, y, sizePt, bold, italic, color, hAlign, vAlign, null);
        }

        private void DrawText(SKCanvas canvas, PageLayout page, SKRect clipRect, string text, float x, float y,
            float sizePt, bool bold, bool italic, SKColor color, TextAlign hAlign, VAlign vAlign, SKColor? outlineColor)
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

                if (outlineColor.HasValue)
                {
                    TryRecordRuns(page, clipRect, fontContext, text, drawX, baseline, paint.Color, true);
                    canvas.Save();
                    canvas.ClipRect(clipRect);
                    DrawRuns(canvas, fontContext, text, drawX, baseline, paint);
                    canvas.Restore();
                }
                else if (!TryRecordRuns(page, clipRect, fontContext, text, drawX, baseline, paint.Color))
                {
                    DrawRuns(canvas, fontContext, text, drawX, baseline, paint);
                }
                DrawTextOutline(canvas, clipRect, fontContext, text, drawX, baseline, fontValue, outlineColor);
            }
        }

        private void DrawTitleText(SKCanvas canvas, PageLayout page, SKRect clipRect, string text, float x, float y,
            FontValue fontValue, SKColor color, TextAlign hAlign, VAlign vAlign, SKColor? outlineColor)
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
                if (outlineColor.HasValue)
                {
                    TryRecordRuns(page, clipRect, fontContext, text, drawX, baseline, paint.Color, true);
                    DrawRuns(canvas, fontContext, text, drawX, baseline, paint);
                }
                else if (!TryRecordRuns(page, clipRect, fontContext, text, drawX, baseline, paint.Color))
                {
                    DrawRuns(canvas, fontContext, text, drawX, baseline, paint);
                }
                DrawTextOutline(canvas, clipRect, fontContext, text, drawX, baseline, fontValue, outlineColor);
                canvas.Restore();
            }
        }

        private void DrawTextOutline(SKCanvas canvas, SKRect clipRect, CellFontContext fontContext, string text,
            float x, float baseline, FontValue fontValue, SKColor? outlineColor)
        {
            if (!outlineColor.HasValue)
            {
                return;
            }

            using (var paint = TextPaint(fontValue, outlineColor.Value))
            {
                paint.Style = SKPaintStyle.Stroke;
                paint.StrokeWidth = 0.75f;
                canvas.Save();
                canvas.ClipRect(clipRect);
                DrawRuns(canvas, fontContext, text, x, baseline, paint);
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
            return TryRecordRuns(page, clipRect, fontContext, text, x, baseline, color, false);
        }

        private bool TryRecordRuns(PageLayout page, SKRect clipRect, CellFontContext fontContext, string text,
            float x, float baseline, SKColor color, bool invisible)
        {
            var session = _context.PdfTextSession;
            if (session == null || page == null || string.IsNullOrEmpty(text))
            {
                return false;
            }

            // The clip can cover a whole chart even for a single-line date label.
            // Its size must not decide whether the text can share font glyphs.
            if (!IsTokenizableChartText(text))
            {
                return false;
            }

            var runs = fontContext.SplitRuns(text);
            var cursor = x;
            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                session.RecordChartRun(page, clipRect, run.Text, run.Typeface, color,
                    fontContext.SizePt, cursor, baseline, invisible);
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
                    PdfTextPathRenderer.DrawText(canvas, run.Text, cursor, baseline, font,
                        paint.Color, paint.Style, paint.StrokeWidth);
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
            return new FontValue { Name = _context.Colors.MinorFontName, Size = sizePt, Bold = bold, Italic = italic };
        }
    }
}
