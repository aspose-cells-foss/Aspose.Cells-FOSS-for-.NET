using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Draws a <see cref="ParsedShape"/> into a rectangle: preset geometries (rect, ellipse,
    /// triangle, block arrows, plus/cross, ...) with fill/line, straight line connectors with
    /// arrowheads, and centered text. Unknown geometries fall back to a rounded/plain rectangle.
    /// </summary>
    internal sealed class ShapeRenderer
    {
        private readonly RenderContext _context;

        public ShapeRenderer(RenderContext context)
        {
            _context = context;
        }

        public void Draw(SKCanvas canvas, SKRect rect, ParsedShape shape)
        {
            Draw(canvas, null, rect, shape, null, null);
        }

        public void Draw(SKCanvas canvas, PageLayout page, SKRect rect, ParsedShape shape)
        {
            Draw(canvas, page, rect, shape, null, null);
        }

        public void Draw(SKCanvas canvas, PageLayout page, SKRect rect, ParsedShape shape, ShapeModel shapeModel, IList<ShapeModel> shapes)
        {
            if (IsConnector(shape.Geometry))
            {
                DrawConnector(canvas, page, rect, shape, shapeModel, shapes);
                return;
            }

            var restore = false;
            if (Math.Abs(shape.RotationDeg) > 0.01d)
            {
                canvas.Save();
                canvas.RotateDegrees((float)shape.RotationDeg, rect.MidX, rect.MidY);
                restore = true;
            }

            DrawFilledShape(canvas, page, rect, shape);

            if (restore)
            {
                canvas.Restore();
            }
        }

        private static bool IsConnector(string geometry)
        {
            return geometry == "line"
                || geometry.StartsWith("straightConnector", StringComparison.Ordinal)
                || geometry.StartsWith("bentConnector", StringComparison.Ordinal)
                || geometry.StartsWith("curvedConnector", StringComparison.Ordinal);
        }

        // "med" arrowhead sizing: ~3x the line width in both length and total width. Excel clamps
        // the scaling width to a minimum so arrowheads stay visible on thin lines (a 0.75pt line
        // still gets a ~7pt head), so scale by max(lineWidth, MinArrowScale).
        private const float MinArrowScalePt = 1.1f;

        private void DrawConnector(SKCanvas canvas, PageLayout page, SKRect rect, ParsedShape shape, ShapeModel shapeModel, IList<ShapeModel> shapes)
        {
            if (!shape.HasLine)
            {
                return;
            }

            if (shape.Geometry.StartsWith("curvedConnector", StringComparison.Ordinal))
            {
                DrawCurvedConnector(canvas, page, rect, shape, shapeModel, shapes);
                return;
            }

            var width = (float)Math.Max(0.5d, shape.LineWidthPt);
            var points = TryBuildConnectedPolyline(page, shapeModel, shapes, shape);
            if (points == null)
            {
                points = BuildConnectorPolyline(rect, shape);
            }

            if (points == null || points.Length < 2)
            {
                return;
            }

            var start = points[0];
            var end = points[points.Length - 1];
            var dx = end.X - start.X;
            var dy = end.Y - start.Y;
            var len = (float)Math.Sqrt(dx * dx + dy * dy);
            if (len < 0.01f)
            {
                return;
            }

            var arrowScale = Math.Max(width, MinArrowScalePt);

            // A filled arrowhead (triangle/stealth/diamond) hides the line end, so stop the line at
            // its base with a flat cap; an open "arrow" (V) sits on top of the line, so the line runs
            // to the tip.
            var drawPoints = CopyPoints(points);
            if (HasArrow(shape.HeadEnd))
            {
                ShortenPolylineStart(drawPoints, FilledArrow(shape.HeadEnd) ? arrowScale * ArrowLineHideMultiplier(shape.HeadEnd) : 0f);
            }

            if (HasArrow(shape.TailEnd))
            {
                ShortenPolylineEnd(drawPoints, FilledArrow(shape.TailEnd) ? arrowScale * ArrowLineHideMultiplier(shape.TailEnd) : 0f);
            }

            using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, Color = shape.LineColor, StrokeWidth = width, IsAntialias = true, StrokeCap = SKStrokeCap.Butt, StrokeJoin = SKStrokeJoin.Miter })
            using (var path = new SKPath())
            {
                path.MoveTo(drawPoints[0]);
                for (var i = 1; i < drawPoints.Length; i++)
                {
                    path.LineTo(drawPoints[i]);
                }

                canvas.DrawPath(path, paint);
            }

            var headDirection = SegmentDirection(points[1], points[0]);
            var tailDirection = SegmentDirection(points[points.Length - 2], points[points.Length - 1]);
            DrawArrowhead(canvas, shape.HeadEnd, start, headDirection.X, headDirection.Y, width, shape.LineColor);
            DrawArrowhead(canvas, shape.TailEnd, end, tailDirection.X, tailDirection.Y, width, shape.LineColor);
        }

        private SKPoint[] TryBuildConnectedPolyline(PageLayout page, ShapeModel shapeModel, IList<ShapeModel> shapes, ParsedShape shape)
        {
            if (page == null || shapeModel == null || shapes == null)
            {
                return null;
            }

            var startShape = FindShapeByDrawingObjectId(shapes, shapeModel.StartConnectionShapeId);
            var endShape = FindShapeByDrawingObjectId(shapes, shapeModel.EndConnectionShapeId);
            if (startShape == null || endShape == null)
            {
                return null;
            }

            var start = ConnectionPoint(page, startShape, shapeModel.StartConnectionSite);
            var end = ConnectionPoint(page, endShape, shapeModel.EndConnectionSite);
            if (shape.Geometry.StartsWith("bentConnector", StringComparison.Ordinal))
            {
                return new[]
                {
                    start,
                    new SKPoint(start.X, end.Y),
                    end,
                };
            }

            if (shape.Geometry.StartsWith("straightConnector", StringComparison.Ordinal) || shape.Geometry == "line")
            {
                return new[]
                {
                    start,
                    end,
                };
            }

            return null;
        }

        private void DrawCurvedConnector(SKCanvas canvas, PageLayout page, SKRect rect, ParsedShape shape, ShapeModel shapeModel, IList<ShapeModel> shapes)
        {
            var width = (float)Math.Max(0.5d, shape.LineWidthPt);
            var arrowScale = Math.Max(width, MinArrowScalePt);

            var start = TransformNormalizedPoint(rect, shape, 0f, 0f);
            var end = TransformNormalizedPoint(rect, shape, 1f, 1f);
            ResolveConnectorEndpoints(page, shapeModel, shapes, ref start, ref end);

            var control1 = new SKPoint(
                start.X + (end.X - start.X) * 0.28f,
                start.Y - rect.Height * 0.12f);
            var control2 = new SKPoint(
                end.X - (end.X - start.X) * 0.18f,
                end.Y + rect.Height * 0.36f);

            var endDirection = SegmentDirection(control2, end);
            var headDirection = SegmentDirection(control1, start);

            var visibleStart = start;
            var visibleEnd = end;
            if (HasArrow(shape.HeadEnd))
            {
                visibleStart = new SKPoint(
                    start.X + headDirection.X * arrowScale * ArrowLineHideMultiplier(shape.HeadEnd),
                    start.Y + headDirection.Y * arrowScale * ArrowLineHideMultiplier(shape.HeadEnd));
            }

            if (HasArrow(shape.TailEnd))
            {
                visibleEnd = new SKPoint(
                    end.X - endDirection.X * arrowScale * ArrowLineHideMultiplier(shape.TailEnd),
                    end.Y - endDirection.Y * arrowScale * ArrowLineHideMultiplier(shape.TailEnd));
            }

            using (var path = new SKPath())
            using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, Color = shape.LineColor, StrokeWidth = width, IsAntialias = true, StrokeCap = SKStrokeCap.Butt, StrokeJoin = SKStrokeJoin.Round })
            {
                path.MoveTo(visibleStart);
                path.CubicTo(control1, control2, visibleEnd);
                canvas.DrawPath(path, paint);
            }

            DrawArrowhead(canvas, shape.HeadEnd, start, headDirection.X, headDirection.Y, width, shape.LineColor);
            DrawArrowhead(canvas, shape.TailEnd, end, endDirection.X, endDirection.Y, width, shape.LineColor);
        }

        private void ResolveConnectorEndpoints(PageLayout page, ShapeModel connector, IList<ShapeModel> shapes, ref SKPoint start, ref SKPoint end)
        {
            if (page == null || connector == null || shapes == null)
            {
                return;
            }

            var startShape = FindShapeByDrawingObjectId(shapes, connector.StartConnectionShapeId);
            if (startShape != null)
            {
                start = ConnectionPoint(page, startShape, connector.StartConnectionSite);
            }

            var endShape = FindShapeByDrawingObjectId(shapes, connector.EndConnectionShapeId);
            if (endShape != null)
            {
                end = ConnectionPoint(page, endShape, connector.EndConnectionSite);
            }
        }

        private static ShapeModel FindShapeByDrawingObjectId(IList<ShapeModel> shapes, int drawingObjectId)
        {
            if (shapes == null || drawingObjectId < 0)
            {
                return null;
            }

            for (var i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                if (shape != null && shape.DrawingObjectId == drawingObjectId)
                {
                    return shape;
                }
            }

            return null;
        }

        private static SKPoint ConnectionPoint(PageLayout page, ShapeModel shape, int site)
        {
            var rect = ShapeRect(page, shape);
            if (site == 2
                && !string.IsNullOrEmpty(shape.RawElementXml)
                && shape.RawElementXml.IndexOf("wordArtVertRtl", StringComparison.Ordinal) >= 0)
            {
                return new SKPoint(rect.Left, rect.MidY);
            }

            if (site == 0)
            {
                return new SKPoint(rect.MidX, rect.Top);
            }

            if (site == 1)
            {
                return new SKPoint(rect.Left, rect.MidY);
            }

            if (site == 2)
            {
                return new SKPoint(rect.MidX, rect.Bottom);
            }

            if (site == 3)
            {
                return new SKPoint(rect.Right, rect.MidY);
            }

            return new SKPoint(rect.MidX, rect.MidY);
        }

        private static SKRect ShapeRect(PageLayout page, ShapeModel shape)
        {
            var layout = page.Sheet;
            var originX = layout.ColumnStartPt[page.StartColumn];
            var originY = layout.RowStartPt[page.StartRow];
            var left = ColumnEdgePt(layout, shape.UpperLeftColumn, shape.UpperLeftColumnOffset) - originX;
            var top = RowEdgePt(layout, shape.UpperLeftRow, shape.UpperLeftRowOffset) - originY;
            var right = ColumnEdgePt(layout, shape.LowerRightColumn, shape.LowerRightColumnOffset) - originX;
            var bottom = RowEdgePt(layout, shape.LowerRightRow, shape.LowerRightRowOffset) - originY;
            return new SKRect((float)left, (float)top, (float)right, (float)bottom);
        }

        private static double ColumnEdgePt(SheetLayout layout, int columnIndex, long offsetEmu)
        {
            return layout.ColumnStartPt[columnIndex] + offsetEmu / 12700d;
        }

        private static double RowEdgePt(SheetLayout layout, int rowIndex, long offsetEmu)
        {
            return layout.RowStartPt[rowIndex] + offsetEmu / 12700d;
        }

        private static bool HasArrow(string type)
        {
            return !string.IsNullOrEmpty(type) && type != "none";
        }

        private static bool FilledArrow(string type)
        {
            return type == "triangle" || type == "stealth" || type == "diamond" || type == "oval";
        }

        private void DrawArrowhead(SKCanvas canvas, string type, SKPoint tip, float dirX, float dirY, float lineWidth, SKColor color)
        {
            if (!HasArrow(type))
            {
                return;
            }

            var px = -dirY;
            var py = dirX;
            var arrowScale = Math.Max(lineWidth, MinArrowScalePt);
            var arrowLen = arrowScale * ArrowLengthMultiplier(type);
            var halfW = arrowScale * ArrowHalfWidthMultiplier(type);
            var baseX = tip.X - dirX * arrowLen;
            var baseY = tip.Y - dirY * arrowLen;

            if (!FilledArrow(type))
            {
                // Open "arrow": two strokes forming a V at the tip.
                using (var paint = new SKPaint { Style = SKPaintStyle.Stroke, Color = color, StrokeWidth = lineWidth, IsAntialias = true, StrokeCap = SKStrokeCap.Round })
                {
                    canvas.DrawLine(tip.X, tip.Y, baseX + px * halfW, baseY + py * halfW, paint);
                    canvas.DrawLine(tip.X, tip.Y, baseX - px * halfW, baseY - py * halfW, paint);
                }
                return;
            }

            using (var path = new SKPath())
            using (var paint = new SKPaint { Style = SKPaintStyle.Fill, Color = color, IsAntialias = true })
            {
                path.MoveTo(tip.X, tip.Y);
                path.LineTo(baseX + px * halfW, baseY + py * halfW);
                path.LineTo(baseX - px * halfW, baseY - py * halfW);
                path.Close();
                canvas.DrawPath(path, paint);
            }
        }

        private static float ArrowLengthMultiplier(string type)
        {
            if (type == "triangle")
            {
                return 0.17f;
            }

            if (type == "stealth")
            {
                return 0.90f;
            }

            return 1.05f;
        }

        private static float ArrowHalfWidthMultiplier(string type)
        {
            if (type == "triangle")
            {
                return 0.92f;
            }

            if (type == "stealth")
            {
                return 1.05f;
            }

            return 1.15f;
        }

        private static float ArrowLineHideMultiplier(string type)
        {
            if (type == "triangle")
            {
                // Excel's filled connector arrowheads are compact wedges that sit on top of the
                // line body rather than replacing a long tail segment.
                return 0.00f;
            }

            return ArrowLengthMultiplier(type);
        }

        private void DrawFilledShape(SKCanvas canvas, PageLayout page, SKRect rect, ParsedShape shape)
        {
            using (var path = BuildGeometry(shape, rect))
            {
                if (shape.HasFill)
                {
                    using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Color = shape.FillColor, IsAntialias = true })
                    {
                        canvas.DrawPath(path, fill);
                    }
                }

                if (shape.HasLine)
                {
                    using (var stroke = new SKPaint { Style = SKPaintStyle.Stroke, Color = shape.LineColor, StrokeWidth = (float)Math.Max(0.5d, shape.LineWidthPt), IsAntialias = true, StrokeJoin = SKStrokeJoin.Miter })
                    {
                        canvas.DrawPath(path, stroke);
                    }
                }
            }

            DrawText(canvas, page, rect, shape);
        }

        private static SKPath BuildGeometry(ParsedShape shape, SKRect r)
        {
            var geometry = shape.Geometry;
            var path = new SKPath();
            float l = r.Left, t = r.Top, right = r.Right, b = r.Bottom, w = r.Width, h = r.Height;
            float cx = r.MidX, cy = r.MidY;

            switch (geometry)
            {
                case "corner":
                    CornerShape(path, r, shape.Adjusts);
                    break;
                case "ellipse":
                    path.AddOval(r);
                    break;
                case "roundRect":
                    path.AddRoundRect(r, Math.Min(w, h) * 0.16f, Math.Min(w, h) * 0.16f);
                    break;
                case "triangle":
                    {
                        // The apex x-position is the adjust value (default 0.5 = isosceles; 1.0 puts
                        // the apex at the right edge, giving a right triangle).
                        var apexAdj = shape.Adjusts != null && shape.Adjusts.Length > 0 ? shape.Adjusts[0] : 0.5;
                        path.MoveTo(l + (float)(w * apexAdj), t);
                        path.LineTo(right, b);
                        path.LineTo(l, b);
                        path.Close();
                    }
                    break;
                case "rtTriangle":
                    path.MoveTo(l, t); path.LineTo(l, b); path.LineTo(right, b); path.Close();
                    break;
                case "diamond":
                    path.MoveTo(cx, t); path.LineTo(right, cy); path.LineTo(cx, b); path.LineTo(l, cy); path.Close();
                    break;
                case "rightArrow":
                    BlockArrowRight(path, r);
                    break;
                case "leftArrow":
                    BlockArrowLeft(path, r);
                    break;
                case "upArrow":
                    BlockArrowUp(path, r);
                    break;
                case "downArrow":
                    BlockArrowDown(path, r);
                    break;
                case "mathPlus":
                case "plus":
                    PlusShape(path, r);
                    break;
                default:
                    path.AddRect(r);
                    break;
            }

            return path;
        }

        private static void CornerShape(SKPath path, SKRect r, double[] adj)
        {
            // L-shape: a vertical arm on the left and a horizontal arm on the bottom; the top-right is
            // cut out. Both arm thicknesses are relative to the shape's smaller dimension (ss), so the
            // arms stay equally thick on a non-square box (per the DrawingML "corner" preset).
            var a1 = adj != null && adj.Length > 0 ? adj[0] : 0.5;
            var a2 = adj != null && adj.Length > 1 ? adj[1] : 0.5;
            var ss = Math.Min(r.Width, r.Height);
            var dx1 = (float)(ss * a2);
            var dy1 = (float)(ss * a1);
            float l = r.Left, t = r.Top, right = r.Right, b = r.Bottom;
            path.MoveTo(l, t);
            path.LineTo(l + dx1, t);
            path.LineTo(l + dx1, b - dy1);
            path.LineTo(right, b - dy1);
            path.LineTo(right, b);
            path.LineTo(l, b);
            path.Close();
        }

        private static void PlusShape(SKPath path, SKRect r)
        {
            // mathPlus preset: arm thickness = ss * adj (default adj 0.2352); the plus arms reach only
            // 73.49% of the box (dx2 = w * 73490/200000), centered - they do NOT fill the box.
            var d = Math.Min(r.Width, r.Height) * 0.2352f;
            var half = d / 2f;
            float cx = r.MidX, cy = r.MidY;
            var hx = r.Width * 0.367450f;   // half horizontal extent of the plus
            var hy = r.Height * 0.367450f;  // half vertical extent of the plus
            float l = cx - hx, right = cx + hx, t = cy - hy, b = cy + hy;

            path.MoveTo(cx - half, t);
            path.LineTo(cx + half, t);
            path.LineTo(cx + half, cy - half);
            path.LineTo(right, cy - half);
            path.LineTo(right, cy + half);
            path.LineTo(cx + half, cy + half);
            path.LineTo(cx + half, b);
            path.LineTo(cx - half, b);
            path.LineTo(cx - half, cy + half);
            path.LineTo(l, cy + half);
            path.LineTo(l, cy - half);
            path.LineTo(cx - half, cy - half);
            path.Close();
        }

        private static void BlockArrowRight(SKPath path, SKRect r)
        {
            float headW = r.Width * 0.10f, shaftH = r.Height * 0.58f, headInset = r.Height * 0.16f;
            float sy0 = r.MidY - shaftH / 2f, sy1 = r.MidY + shaftH / 2f, hx = r.Right - headW;
            path.MoveTo(r.Left, sy0);
            path.LineTo(hx, sy0); path.LineTo(hx, r.Top + headInset); path.LineTo(r.Right, r.MidY);
            path.LineTo(hx, r.Bottom - headInset); path.LineTo(hx, sy1); path.LineTo(r.Left, sy1);
            path.Close();
        }

        private static void BlockArrowLeft(SKPath path, SKRect r)
        {
            float headW = r.Width * 0.4f, shaftH = r.Height * 0.5f;
            float sy0 = r.MidY - shaftH / 2f, sy1 = r.MidY + shaftH / 2f, hx = r.Left + headW;
            path.MoveTo(r.Right, sy0);
            path.LineTo(hx, sy0); path.LineTo(hx, r.Top); path.LineTo(r.Left, r.MidY);
            path.LineTo(hx, r.Bottom); path.LineTo(hx, sy1); path.LineTo(r.Right, sy1);
            path.Close();
        }

        private static void BlockArrowUp(SKPath path, SKRect r)
        {
            float headH = r.Height * 0.4f, shaftW = r.Width * 0.5f;
            float sx0 = r.MidX - shaftW / 2f, sx1 = r.MidX + shaftW / 2f, hy = r.Top + headH;
            path.MoveTo(sx0, r.Bottom);
            path.LineTo(sx0, hy); path.LineTo(r.Left, hy); path.LineTo(r.MidX, r.Top);
            path.LineTo(r.Right, hy); path.LineTo(sx1, hy); path.LineTo(sx1, r.Bottom);
            path.Close();
        }

        private static void BlockArrowDown(SKPath path, SKRect r)
        {
            float headH = r.Height * 0.4f, shaftW = r.Width * 0.5f;
            float sx0 = r.MidX - shaftW / 2f, sx1 = r.MidX + shaftW / 2f, hy = r.Bottom - headH;
            path.MoveTo(sx0, r.Top);
            path.LineTo(sx0, hy); path.LineTo(r.Left, hy); path.LineTo(r.MidX, r.Bottom);
            path.LineTo(r.Right, hy); path.LineTo(sx1, hy); path.LineTo(sx1, r.Top);
            path.Close();
        }

        private void DrawText(SKCanvas canvas, PageLayout page, SKRect rect, ParsedShape shape)
        {
            if (string.IsNullOrEmpty(shape.Text))
            {
                return;
            }

            var font = new FontValue { Name = "Calibri", Size = shape.TextSizePt, Bold = shape.TextBold };
            var hAlign = shape.TextHAlign == "l" ? SKTextAlign.Left : shape.TextHAlign == "r" ? SKTextAlign.Right : SKTextAlign.Center;
            var fontContext = _context.GetFontContext(font);
            using (var paint = new SKPaint
            {
                Color = shape.TextColor,
                IsAntialias = true,
                Typeface = _context.Fonts.Resolve(font),
                TextSize = (float)shape.TextSizePt,
                TextAlign = hAlign,
            })
            {
                var metrics = paint.FontMetrics;

                float x;
                if (hAlign == SKTextAlign.Left) x = rect.Left + (float)shape.TextInsetLeftPt;
                else if (hAlign == SKTextAlign.Right) x = rect.Right - (float)shape.TextInsetLeftPt;
                else x = rect.MidX;

                var width = fontContext.Measure(shape.Text);
                var drawX = hAlign == SKTextAlign.Left ? x : hAlign == SKTextAlign.Center ? x - width / 2f : x - width;

                // Vertical anchor: top / center / bottom.
                float baseline;
                if (shape.TextVAnchor == "t") baseline = rect.Top + (float)shape.TextInsetTopPt - metrics.Ascent;
                else if (shape.TextVAnchor == "b") baseline = rect.Bottom - (float)shape.TextInsetTopPt - metrics.Descent;
                else baseline = rect.MidY - (metrics.Ascent + metrics.Descent) / 2f;

                if (IsVerticalText(shape))
                {
                    DrawVerticalText(canvas, rect, shape, fontContext, paint);
                    return;
                }

                if (!TryRecordText(page, rect, fontContext, shape.Text, paint.Color, paint.TextSize, drawX, baseline, shape.RotationDeg))
                {
                    using (var fontObject = new SKFont(paint.Typeface, paint.TextSize))
                    {
                        fontObject.Subpixel = true;
                        PdfTextPathRenderer.DrawText(canvas, shape.Text, drawX, baseline, fontObject, paint.Color);
                    }
                }
            }
        }

        private bool TryRecordText(PageLayout page, SKRect clipRect, CellFontContext fontContext, string text, SKColor color, float fontSizePt, float x, float baseline, double rotationDeg)
        {
            var session = _context.PdfTextSession;
            if (session == null || page == null || fontContext == null || string.IsNullOrEmpty(text))
            {
                return false;
            }

            // Shape text is already laid out in the local drawing canvas. Re-emitting it through the
            // page-level Type3 text rewrite has produced upside-down text in some text boxes
            // (shape1.xlsx "Start"), so keep shape text on the original path-rendering route.
            return false;
        }

        private static SKPoint[] BuildConnectorPolyline(SKRect rect, ParsedShape shape)
        {
            SKPoint[] points;
            if (shape.Geometry.StartsWith("bentConnector", StringComparison.Ordinal))
            {
                if (Math.Abs(shape.RotationDeg - 270d) <= 0.01d && shape.FlipH && !shape.FlipV)
                {
                    points = new[]
                    {
                        new SKPoint(rect.Left, rect.Top),
                        new SKPoint(rect.Left, rect.Bottom),
                        new SKPoint(rect.Right, rect.Bottom),
                    };
                    return points;
                }

                points = new[]
                {
                    new SKPoint(rect.Left, rect.MidY),
                    new SKPoint(rect.Right, rect.MidY),
                    new SKPoint(rect.Right, rect.Bottom),
                };
            }
            else if (shape.Geometry.StartsWith("curvedConnector", StringComparison.Ordinal))
            {
                points = new[]
                {
                    new SKPoint(rect.Left, rect.MidY),
                    new SKPoint(rect.Left + rect.Width * 0.35f, rect.MidY),
                    new SKPoint(rect.Left + rect.Width * 0.75f, rect.Bottom),
                    new SKPoint(rect.Right, rect.Bottom),
                };
            }
            else
            {
                points = new[]
                {
                    new SKPoint(rect.Left, rect.Top),
                    new SKPoint(rect.Right, rect.Bottom),
                };
            }

            return TransformConnectorPoints(points, rect, shape);
        }

        private static SKPoint[] TransformConnectorPoints(SKPoint[] points, SKRect rect, ParsedShape shape)
        {
            var result = new SKPoint[points.Length];
            for (var i = 0; i < points.Length; i++)
            {
                var point = points[i];
                result[i] = TransformNormalizedPoint(
                    rect,
                    shape,
                    (point.X - rect.Left) / rect.Width,
                    (point.Y - rect.Top) / rect.Height);
            }

            return result;
        }

        private static SKPoint TransformNormalizedPoint(SKRect rect, ParsedShape shape, float x, float y)
        {
            if (shape.FlipH)
            {
                x = 1f - x;
            }

            if (shape.FlipV)
            {
                y = 1f - y;
            }

            var transformed = RotateNormalizedPoint(x, y, shape.RotationDeg);
            return new SKPoint(
                rect.Left + transformed.X * rect.Width,
                rect.Top + transformed.Y * rect.Height);
        }

        private static SKPoint RotateNormalizedPoint(float x, float y, double rotationDeg)
        {
            if (Math.Abs(rotationDeg) <= 0.01d)
            {
                return new SKPoint(x, y);
            }

            var radians = -rotationDeg * Math.PI / 180d;
            var dx = x - 0.5d;
            var dy = y - 0.5d;
            var rotatedX = dx * Math.Cos(radians) - dy * Math.Sin(radians);
            var rotatedY = dx * Math.Sin(radians) + dy * Math.Cos(radians);
            return new SKPoint((float)(0.5d + rotatedX), (float)(0.5d + rotatedY));
        }

        private static SKPoint[] CopyPoints(SKPoint[] points)
        {
            var copy = new SKPoint[points.Length];
            for (var i = 0; i < points.Length; i++)
            {
                copy[i] = points[i];
            }

            return copy;
        }

        private static void ShortenPolylineStart(SKPoint[] points, float shortenBy)
        {
            if (shortenBy <= 0f || points == null || points.Length < 2)
            {
                return;
            }

            var direction = SegmentDirection(points[1], points[0]);
            points[0] = new SKPoint(points[0].X + direction.X * shortenBy, points[0].Y + direction.Y * shortenBy);
        }

        private static void ShortenPolylineEnd(SKPoint[] points, float shortenBy)
        {
            if (shortenBy <= 0f || points == null || points.Length < 2)
            {
                return;
            }

            var last = points.Length - 1;
            var direction = SegmentDirection(points[last - 1], points[last]);
            points[last] = new SKPoint(points[last].X - direction.X * shortenBy, points[last].Y - direction.Y * shortenBy);
        }

        private static SKPoint SegmentDirection(SKPoint from, SKPoint to)
        {
            var dx = to.X - from.X;
            var dy = to.Y - from.Y;
            var length = (float)Math.Sqrt(dx * dx + dy * dy);
            if (length <= 0.01f)
            {
                return new SKPoint(1f, 0f);
            }

            return new SKPoint(dx / length, dy / length);
        }

        private static bool IsVerticalText(ParsedShape shape)
        {
            if (shape == null || string.IsNullOrEmpty(shape.TextVerticalType))
            {
                return false;
            }

            return shape.TextVerticalType == "vert"
                || shape.TextVerticalType == "vert270"
                || shape.TextVerticalType == "wordArtVert"
                || shape.TextVerticalType == "wordArtVertRtl";
        }

        private void DrawVerticalText(SKCanvas canvas, SKRect rect, ParsedShape shape, CellFontContext fontContext, SKPaint paint)
        {
            if (shape.TextVerticalType == "wordArtVert" || shape.TextVerticalType == "wordArtVertRtl")
            {
                DrawWordArtVerticalText(canvas, rect, shape, fontContext, paint);
                return;
            }

            var clockwise = shape.TextVerticalType == "vert270" || shape.TextVerticalType == "wordArtVertRtl";
            canvas.Save();
            canvas.Translate(rect.MidX, rect.MidY);
            canvas.RotateDegrees(clockwise ? 90f : -90f);

            var metrics = paint.FontMetrics;
            var width = fontContext.Measure(shape.Text);
            var drawX = -width / 2f;
            var baseline = -(metrics.Ascent + metrics.Descent) / 2f;
            using (var fontObject = new SKFont(paint.Typeface, paint.TextSize))
            {
                fontObject.Subpixel = true;
                PdfTextPathRenderer.DrawText(canvas, shape.Text, drawX, baseline, fontObject, paint.Color);
            }

            canvas.Restore();
        }

        private void DrawWordArtVerticalText(SKCanvas canvas, SKRect rect, ParsedShape shape, CellFontContext fontContext, SKPaint paint)
        {
            if (string.IsNullOrEmpty(shape.Text))
            {
                return;
            }

            var metrics = paint.FontMetrics;
            var centerX = rect.MidX;
            var charAdvance = Math.Max((float)shape.TextSizePt, metrics.Descent - metrics.Ascent);
            var totalHeight = charAdvance * shape.Text.Length;
            var currentY = rect.MidY - totalHeight * 0.5f;

            canvas.Save();
            canvas.ClipRect(rect);
            using (var fontObject = new SKFont(paint.Typeface, paint.TextSize))
            {
                fontObject.Subpixel = true;
                for (var i = 0; i < shape.Text.Length; i++)
                {
                    var ch = shape.Text[i].ToString();
                    var width = fontContext.Measure(ch);
                    var drawX = centerX - width * 0.5f;
                    var baseline = currentY - metrics.Ascent;
                    PdfTextPathRenderer.DrawText(canvas, ch, drawX, baseline, fontObject, paint.Color);
                    currentY += charAdvance;
                }
            }
            canvas.Restore();
        }
    }
}
