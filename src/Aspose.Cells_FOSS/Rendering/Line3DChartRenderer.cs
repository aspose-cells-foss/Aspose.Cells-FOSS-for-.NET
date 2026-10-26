using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Draws line3DChart series as filled ribbons on a projected value/category/depth box.
    /// The noFill line in a 3D series suppresses its outline, not the filled ribbon.
    /// </summary>
    internal sealed class Line3DChartRenderer
    {
        private const float LabelSizePt = 9f;
        private readonly RenderContext _context;
        private readonly DateSystem _dateSystem;
        private readonly CultureInfo _culture;

        public Line3DChartRenderer(RenderContext context, DateSystem dateSystem, CultureInfo culture)
        {
            _context = context;
            _dateSystem = dateSystem;
            _culture = culture ?? CultureInfo.CurrentCulture;
        }

        public void Draw(SKCanvas canvas, PageLayout page, SKRect area, SKRect plot, ParsedChart chart,
            double min, double max, double major)
        {
            if (canvas == null || chart == null || chart.Series.Count == 0 || max <= min)
            {
                return;
            }

            var projection = new Line3DProjection(plot, chart, PointCount(chart));
            DrawBox(canvas, projection, chart, min, max, major);
            DrawRibbons(canvas, projection, chart, min, max);
            DrawValueLabels(canvas, page, area, projection, chart, min, max, major);
            DrawCategoryLabels(canvas, page, area, projection, chart);
            DrawSeriesLabels(canvas, page, area, projection, chart);
        }

        private void DrawBox(SKCanvas canvas, Line3DProjection projection, ParsedChart chart,
            double min, double max, double major)
        {
            using (var grid = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.ValueGridlineColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.ValueGridlineWidthPt),
                IsAntialias = true
            })
            using (var axis = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = chart.CategoryAxisLineColor,
                StrokeWidth = (float)Math.Max(0.5d, chart.CategoryAxisLineWidthPt),
                IsAntialias = true
            })
            {
                if (chart.HasValueGridlines && major > 0d)
                {
                    var steps = (int)Math.Round((max - min) / major);
                    for (var i = 0; i <= steps; i++)
                    {
                        var value = Math.Min(max, min + i * major);
                        var y = (value - min) / (max - min);
                        var front = projection.Project(0d, y, 0d);
                        var backLeft = projection.Project(0d, y, 1d);
                        var backRight = projection.Project(1d, y, 1d);
                        canvas.DrawLine(front, backLeft, grid);
                        canvas.DrawLine(backLeft, backRight, grid);
                    }
                }

                var frontLeft = projection.Project(0d, 0d, 0d);
                var frontRight = projection.Project(1d, 0d, 0d);
                if (chart.HasCategoryAxisLine)
                {
                    canvas.DrawLine(frontLeft, frontRight, axis);
                }

                var monthlyTicks = Area3DDateAxisMath.BuildMonthlyTicks(chart, _dateSystem, _culture);
                var count = monthlyTicks.Count > 0 ? monthlyTicks.Count : PointCount(chart);
                var between = !chart.HasDateCategoryAxis && chart.CrossBetween;
                for (var i = 0; i < count + (between ? 1 : 0); i++)
                {
                    var x = monthlyTicks.Count > 0
                        ? monthlyTicks[i].Item1
                        : between ? (double)i / Math.Max(1, count)
                        : Area3DDateAxisMath.CategoryPosition(chart, i, count, _dateSystem);
                    var tick = projection.Project(x, 0d, 0d);
                    canvas.DrawLine(tick.X, tick.Y, tick.X, tick.Y + 2.5f, axis);
                }
            }
        }

        private void DrawRibbons(SKCanvas canvas, Line3DProjection projection, ParsedChart chart, double min, double max)
        {
            var count = PointCount(chart);
            for (var seriesIndex = chart.Series.Count - 1; seriesIndex >= 0; seriesIndex--)
            {
                var series = chart.Series[seriesIndex];
                var values = ChartStackingMath.BuildDisplayValues(chart, seriesIndex);
                var points = new List<SKPoint>();
                for (var i = 0; i < values.Count; i++)
                {
                    if (!values[i].HasValue)
                    {
                        continue;
                    }

                    var x = Area3DDateAxisMath.CategoryPosition(chart, i, count, _dateSystem);
                    var y = Math.Max(0d, Math.Min(1d, (values[i].Value - min) / (max - min)));
                    points.Add(new SKPoint((float)x, (float)y));
                }

                if (points.Count < 2)
                {
                    continue;
                }

                for (var i = 1; i < points.Count; i++)
                {
                    var point = points[i];
                    var j = i - 1;
                    while (j >= 0 && points[j].X > point.X)
                    {
                        points[j + 1] = points[j];
                        j--;
                    }
                    points[j + 1] = point;
                }

                using (var ribbon = new SKPath())
                using (var fill = new SKPaint
                {
                    Style = SKPaintStyle.Fill,
                    IsAntialias = true
                })
                {
                    var faces = new List<SKPoint[]>();
                    for (var i = 1; i < points.Count; i++)
                    {
                        faces.Add(projection.RibbonCorners(points[i - 1], points[i], seriesIndex));
                    }

                    // Underpaint the continuous silhouette so antialiasing on
                    // shared face edges cannot leave a white crack at a bend.
                    ribbon.MoveTo(faces[0][0]);
                    for (var i = 0; i < faces.Count; i++)
                    {
                        ribbon.LineTo(faces[i][1]);
                    }
                    for (var i = faces.Count - 1; i >= 0; i--)
                    {
                        ribbon.LineTo(faces[i][2]);
                    }
                    ribbon.LineTo(faces[0][3]);
                    ribbon.Close();
                    fill.Color = Shade(series.Color, 0.76d);
                    if (series.Color.Alpha == 255)
                    {
                        canvas.DrawPath(ribbon, fill);
                    }

                    for (var i = 0; i < faces.Count; i++)
                    {
                        var corners = faces[i];
                        ribbon.Reset();
                        ribbon.MoveTo(corners[0]);
                        ribbon.LineTo(corners[1]);
                        ribbon.LineTo(corners[2]);
                        ribbon.LineTo(corners[3]);
                        ribbon.Close();
                        fill.Color = Shade(series.Color, projection.RibbonShade(points[i], points[i + 1], corners));
                        canvas.DrawPath(ribbon, fill);
                    }
                }
            }
        }

        private void DrawValueLabels(SKCanvas canvas, PageLayout page, SKRect area, Line3DProjection projection, ParsedChart chart,
            double min, double max, double major)
        {
            if (major <= 0d)
            {
                return;
            }

            var steps = (int)Math.Round((max - min) / major);
            for (var i = 0; i <= steps; i++)
            {
                var value = Math.Min(max, min + i * major);
                var point = projection.Project(0d, (value - min) / (max - min), 0d);
                var label = ChartXmlParser.FormatLabel(value, chart.ValueFormatCode, _dateSystem, _culture);
                DrawText(canvas, page, area, label != null ? label.TrimEnd() : string.Empty,
                    point.X - 8f, point.Y, chart.AxisLabelColor, true, false);
            }
        }

        private void DrawCategoryLabels(SKCanvas canvas, PageLayout page, SKRect area, Line3DProjection projection, ParsedChart chart)
        {
            var ticks = Area3DDateAxisMath.BuildMonthlyTicks(chart, _dateSystem, _culture);
            if (ticks.Count == 0)
            {
                var count = PointCount(chart);
                for (var i = 0; i < count; i++)
                {
                    var label = i < chart.Categories.Count ? chart.Categories[i] : string.Empty;
                    if (string.IsNullOrEmpty(label) && i < chart.CategoryValues.Count
                        && chart.CategoryValues[i].HasValue)
                    {
                        var format = string.IsNullOrEmpty(chart.CategoryFormatCode)
                            ? "mmm-yy" : chart.CategoryFormatCode;
                        label = ChartXmlParser.FormatLabel(chart.CategoryValues[i].Value,
                            format, _dateSystem, _culture);
                    }

                    var x = Area3DDateAxisMath.CategoryPosition(chart, i, count, _dateSystem);
                    var point = projection.Project(x, 0d, 0d);
                    DrawText(canvas, page, area, label, point.X, point.Y + 7f,
                        chart.AxisLabelColor, false, true);
                }

                return;
            }

            for (var i = 0; i < ticks.Count; i++)
            {
                var point = projection.Project(ticks[i].Item1, 0d, 0d);
                DrawText(canvas, page, area, ticks[i].Item2, point.X - 4f, point.Y,
                    chart.AxisLabelColor, true, false,
                    (float)chart.CategoryAxisTextRotationDeg, point.X, point.Y);
            }
        }

        private void DrawSeriesLabels(SKCanvas canvas, PageLayout page, SKRect area, Line3DProjection projection, ParsedChart chart)
        {
            if (!chart.HasVisibleSeriesAxis)
            {
                return;
            }

            for (var i = chart.Series.Count - 1; i >= 0; i--)
            {
                var depth = projection.SeriesDepth(i);
                var point = projection.Project(1d, 0d, depth);
                DrawText(canvas, page, area, chart.Series[i].Name, point.X + 5f, point.Y,
                    chart.SeriesAxisLabelColor, false, false);
            }
        }

        private void DrawText(SKCanvas canvas, PageLayout page, SKRect clip, string text, float x, float y,
            SKColor color, bool alignRight, bool alignTop,
            float rotationDeg = 0f, float originX = 0f, float originY = 0f)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var fontValue = new FontValue { Name = "Calibri", Size = LabelSizePt };
            var fontContext = _context.GetFontContext(fontValue);
            var drawX = alignRight ? x - fontContext.Measure(text)
                : alignTop ? x - fontContext.Measure(text) * 0.5f : x;
            using (var paint = new SKPaint
            {
                Typeface = _context.Fonts.Resolve(fontValue),
                TextSize = LabelSizePt,
                IsAntialias = true
            })
            {
                var metrics = paint.FontMetrics;
                var baseline = alignTop ? y - metrics.Ascent : y - (metrics.Ascent + metrics.Descent) * 0.5f;
                var runs = fontContext.SplitRuns(text);
                var cursor = drawX;
                var session = _context.PdfTextSession;
                if (session != null && page != null && text.IndexOf('\r') < 0 && text.IndexOf('\n') < 0)
                {
                    for (var i = 0; i < runs.Count; i++)
                    {
                        var run = runs[i];
                        if (rotationDeg == 0f)
                        {
                            session.RecordChartRun(page, clip, run.Text, run.Typeface, color,
                                LabelSizePt, cursor, baseline);
                        }
                        else
                        {
                            // The PDF session applies the page transform and rotation pivot;
                            // only the glyph position is relative to that pivot.
                            session.RecordChartRotatedRun(page, clip, run.Text, run.Typeface, color,
                                LabelSizePt, originX, originY, cursor - originX, baseline - originY, rotationDeg);
                        }
                        cursor += run.WidthPt;
                    }
                    return;
                }

                canvas.Save();
                canvas.ClipRect(clip);
                canvas.RotateDegrees(rotationDeg, originX, originY);
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

        private static int PointCount(ParsedChart chart)
        {
            var count = Math.Max(chart.Categories.Count, chart.CategoryValues.Count);
            for (var i = 0; i < chart.Series.Count; i++)
            {
                count = Math.Max(count, chart.Series[i].Values.Count);
            }

            return count;
        }

        private static SKColor Shade(SKColor color, double factor)
        {
            return new SKColor((byte)Math.Round(color.Red * factor),
                (byte)Math.Round(color.Green * factor),
                (byte)Math.Round(color.Blue * factor), color.Alpha);
        }
    }
}
