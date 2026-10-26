using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class Column3DSceneRenderer
    {
        private readonly RenderContext _context;
        private readonly DateSystem _dateSystem;
        private readonly CultureInfo _culture;

        public Column3DSceneRenderer(RenderContext context, DateSystem dateSystem, CultureInfo culture)
        {
            _context = context;
            _dateSystem = dateSystem;
            _culture = culture;
        }

        public void Draw(SKCanvas canvas, SKRect area, ParsedChart chart,
            double min, double max, double major)
        {
            var count = chart.Categories.Count;
            for (var i = 0; i < chart.Series.Count; i++)
            {
                count = Math.Max(count, chart.Series[i].Values.Count);
            }
            if (count == 0)
            {
                return;
            }
            var categories = new Column3DCategoryLayout(chart, count, _dateSystem, _culture);
            var rotation = ChartDateAxisLayout.Resolve(chart, _dateSystem, area.Width * 0.69f).RotationDeg;
            var labelReserve = Math.Abs(rotation) > 45f ? area.Height * 0.25f : 0f;
            // Empty date slots change the column width, not the series-depth scale.
            var projection = new Column3DProjection(area, chart, count, labelReserve);
            if (!chart.MajorUnit.HasValue && major > 0d)
            {
                var axisHeight = Math.Abs(projection.Project(0, 1, 0).Y
                    - projection.Project(0, 0, 0).Y);
                while ((max - min) / major > 1d && axisHeight * major / (max - min) < 18f)
                {
                    major *= 2d;
                }
            }
            canvas.Save();
            canvas.ClipRect(area);
            DrawWalls(canvas, projection, chart, min, max, major, categories.SlotCount);
            DrawColumns(canvas, projection, chart, min, max, categories);
            DrawLabels(canvas, projection, chart, min, max, major, categories, area.Width);
            canvas.Restore();
        }

        private static void DrawWalls(SKCanvas canvas, Column3DProjection p,
            ParsedChart chart, double min, double max, double major, int count)
        {
            DrawWallTexture(canvas, p, true, chart.SideWallImageData);
            DrawWallTexture(canvas, p, false, chart.BackWallImageData);
            using (var grid = new SKPaint
            {
                Color = new SKColor(0xD9, 0xD9, 0xD9), StrokeWidth = 0.6f,
                Style = SKPaintStyle.Stroke, IsAntialias = true
            })
            {
                if (chart.HasMinorValueGridlines)
                {
                    var minor = chart.MinorUnit.GetValueOrDefault(major / 5d);
                    if (minor > 0d)
                    {
                        grid.Color = chart.MinorValueGridlineColor;
                        grid.StrokeWidth = (float)chart.MinorValueGridlineWidthPt;
                        DrawGridLevels(canvas, p, grid, min, max, minor);
                    }
                }
                grid.Color = new SKColor(0xD9, 0xD9, 0xD9);
                grid.StrokeWidth = 0.6f;
                if (chart.HasValueGridlines)
                {
                    DrawGridLevels(canvas, p, grid, min, max, major);
                }
                canvas.DrawLine(p.Project(0, 0, 1), p.Project(0, 0, 0), grid);
                canvas.DrawLine(p.Project(0, 0, 0), p.Project(1, 0, 0), grid);
                for (var i = 0; i <= count; i++)
                {
                    var point = p.Project(i / (double)count, 0, 0);
                    canvas.DrawLine(point.X, point.Y, point.X, point.Y + 2.8f, grid);
                }
                if (chart.HasValueAxisLine)
                {
                    grid.Color = chart.ValueAxisLineColor;
                    grid.StrokeWidth = (float)chart.ValueAxisLineWidthPt;
                    canvas.DrawLine(p.Project(0, 0, 0), p.Project(0, 1, 0), grid);
                    if (string.Equals(chart.ValueAxisMajorTickMark, "cross", StringComparison.Ordinal))
                    {
                        for (var value = min; value <= max + major * 0.001d; value += major)
                        {
                            var point = p.Project(0, (value - min) / (max - min), 0);
                            canvas.DrawLine(point.X - 3f, point.Y, point.X + 3f, point.Y, grid);
                        }
                    }
                }
            }
        }

        private static void DrawWallTexture(SKCanvas canvas, Column3DProjection p,
            bool side, byte[] imageData)
        {
            if (imageData == null || imageData.Length == 0)
            {
                return;
            }
            using (var data = SKData.CreateCopy(imageData))
            using (var image = SKImage.FromEncodedData(data))
            using (var paint = new SKPaint { IsAntialias = true, FilterQuality = SKFilterQuality.High })
            {
                if (image == null)
                {
                    return;
                }
                var size = p.WallSize(side);
                var matrix = p.WallTransform(side);
                canvas.Save();
                canvas.Concat(ref matrix);
                using (var shader = image.ToShader(SKShaderTileMode.Repeat,
                    SKShaderTileMode.Repeat, SKMatrix.CreateScale(0.75f, 0.75f)))
                {
                    paint.Shader = shader;
                    canvas.DrawRect(new SKRect(0, 0, size.Width, size.Height), paint);
                    paint.Shader = null;
                }
                if (side)
                {
                    paint.Color = new SKColor(0, 0, 0, 71);
                    canvas.DrawRect(new SKRect(0, 0, size.Width, size.Height), paint);
                }
                canvas.Restore();
            }
        }

        private static void DrawGridLevels(SKCanvas canvas, Column3DProjection p,
            SKPaint paint, double min, double max, double interval)
        {
            if (interval <= 0d)
            {
                return;
            }
            var count = Math.Min(1000, (int)Math.Round((max - min) / interval));
            for (var i = 0; i <= count; i++)
            {
                var y = i * interval / (max - min);
                var corner = p.Project(0, y, 1);
                canvas.DrawLine(p.Project(0, y, 0), corner, paint);
                canvas.DrawLine(corner, p.Project(1, y, 1), paint);
            }
        }

        private static void DrawColumns(SKCanvas canvas, Column3DProjection p,
            ParsedChart chart, double min, double max, Column3DCategoryLayout categories)
        {
            var count = categories.Centers.Count;
            var width = Column3DProjection.ColumnWidth(chart, categories.SlotCount);
            var depth = Column3DProjection.ColumnDepth(chart);
            var baseline = Math.Max(0d, Math.Min(1d, -min / (max - min)));
            var order = new List<int>();
            for (var s = 0; s < chart.Series.Count; s++)
            {
                for (var c = 0; c < count; c++)
                {
                    if (c < chart.Series[s].Values.Count && chart.Series[s].Values[c].HasValue)
                    {
                        order.Add(s * count + c);
                    }
                }
            }
            order.Sort(delegate(int a, int b)
            {
                var da = p.CameraDepth(categories.Centers[a % count], 0,
                    (a / count + 0.5d) / chart.Series.Count);
                var db = p.CameraDepth(categories.Centers[b % count], 0,
                    (b / count + 0.5d) / chart.Series.Count);
                var comparison = db.CompareTo(da);
                return comparison != 0 ? comparison : a.CompareTo(b);
            });
            for (var i = 0; i < order.Count; i++)
            {
                var s = order[i] / count;
                var c = order[i] % count;
                var series = chart.Series[s];
                var value = Math.Max(0d, Math.Min(1d,
                    (series.Values[c].Value - min) / (max - min)));
                var points = p.ColumnCorners(categories.Centers[c],
                    (s + 0.5d) / chart.Series.Count, width, depth,
                    Math.Min(baseline, value), Math.Max(baseline, value));
                DrawFace(canvas, points, 0, 3, 7, 4, series, 0.72d);
                DrawFace(canvas, points, 3, 2, 6, 7, series, 0.8d);
                DrawFace(canvas, points, 1, 5, 6, 2, series, 0.66d);
                DrawFace(canvas, points, 4, 7, 6, 5, series, 0.85d);
                DrawFace(canvas, points, 0, 4, 5, 1, series, 1d);
            }
        }

        private static void DrawFace(SKCanvas canvas, SKPoint[] points,
            int a, int b, int c, int d, ChartSeries series, double brightness)
        {
            // Outward face winding survives the projection; hidden faces are culled.
            var cross = (points[b].X - points[a].X) * (points[c].Y - points[a].Y)
                - (points[b].Y - points[a].Y) * (points[c].X - points[a].X);
            if (cross <= 0f)
            {
                return;
            }
            using (var path = new SKPath())
            using (var paint = new SKPaint { IsAntialias = true, Style = SKPaintStyle.Fill })
            {
                var color = series.Color;
                paint.Color = new SKColor((byte)Math.Round(color.Red * brightness),
                    (byte)Math.Round(color.Green * brightness),
                    (byte)Math.Round(color.Blue * brightness), color.Alpha);
                path.MoveTo(points[a]);
                path.LineTo(points[b]);
                path.LineTo(points[c]);
                path.LineTo(points[d]);
                path.Close();
                canvas.DrawPath(path, paint);
                if (series.HasVisibleLine)
                {
                    paint.Style = SKPaintStyle.Stroke;
                    paint.Color = series.LineColor;
                    paint.StrokeWidth = (float)series.LineWidthPt;
                    canvas.DrawPath(path, paint);
                }
            }
        }

        private void DrawLabels(SKCanvas canvas, Column3DProjection p, ParsedChart chart,
            double min, double max, double major, Column3DCategoryLayout categories, float width)
        {
            if (major > 0d)
            {
                for (var value = min; value <= max + major * 0.001d; value += major)
                {
                    var point = p.Project(0, (value - min) / (max - min), 0);
                    var label = ChartXmlParser.FormatLabel(value, chart.ValueFormatCode,
                        _dateSystem, _culture);
                    DrawText(canvas, label.TrimEnd(), point.X - 9f, point.Y,
                        chart.AxisLabelColor, 1f);
                }
            }
            var rotation = ChartDateAxisLayout.Resolve(chart, _dateSystem, width * 0.69f).RotationDeg;
            for (var c = 0; c < categories.SlotCount; c++)
            {
                var point = p.Project((c + 0.5d) / categories.SlotCount, 0, 0);
                if (Math.Abs(rotation) > 0.1f)
                {
                    canvas.Save();
                    canvas.Translate(point.X, point.Y + 7f);
                    canvas.RotateDegrees(rotation);
                    DrawText(canvas, categories.Labels[c], 0, 0, chart.AxisLabelColor,
                        rotation < 0f ? 1f : 0f);
                    canvas.Restore();
                }
                else
                {
                    DrawText(canvas, categories.Labels[c], point.X, point.Y + 13f,
                        chart.AxisLabelColor, 0.5f);
                }
            }
            for (var s = 0; s < chart.Series.Count; s++)
            {
                var series = chart.Series[s];
                var z = (s + 0.5d) / chart.Series.Count;
                if (chart.HasVisibleSeriesAxis)
                {
                    var point = p.Project(1, 0, z);
                    DrawText(canvas, series.Name, point.X + 8f, point.Y,
                        chart.SeriesAxisLabelColor, 0f);
                }
                if (!series.ShowValueDataLabels)
                {
                    continue;
                }
                for (var c = 0; c < series.Values.Count; c++)
                {
                    if (!series.Values[c].HasValue)
                    {
                        continue;
                    }
                    var value = series.Values[c].Value;
                    var y = Math.Max(0d, Math.Min(1d, (value - min) / (max - min)));
                    var centered = string.Equals(series.DataLabelPosition, "ctr", StringComparison.Ordinal);
                    var point = p.Project(categories.Centers[c], centered ? y * 0.5d : y, z);
                    var format = string.IsNullOrEmpty(series.ValueFormatCode)
                        ? chart.ValueFormatCode : series.ValueFormatCode;
                    var label = ChartXmlParser.FormatLabel(value, format, _dateSystem, _culture);
                    DrawText(canvas, label.TrimEnd(), point.X, point.Y - (centered ? 0f : 6f),
                        series.DataLabelColor, 0.5f);
                }
            }
        }

        private void DrawText(SKCanvas canvas, string text, float x, float y,
            SKColor color, float alignment)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }
            var fontContext = _context.GetFontContext(new FontValue { Name = "Calibri", Size = 9d });
            var runs = fontContext.SplitRuns(text);
            if (runs.Count == 0)
            {
                return;
            }
            var cursor = x - fontContext.Measure(text) * alignment;
            using (var metricsFont = new SKFont(runs[0].Typeface, 9f))
            {
                var baseline = y - (metricsFont.Metrics.Ascent + metricsFont.Metrics.Descent) * 0.5f;
                for (var i = 0; i < runs.Count; i++)
                {
                    using (var font = new SKFont(runs[i].Typeface, 9f))
                    {
                        font.Subpixel = true;
                        PdfTextPathRenderer.DrawText(canvas, runs[i].Text, cursor, baseline, font, color);
                    }
                    cursor += runs[i].WidthPt;
                }
            }
        }
    }
}
