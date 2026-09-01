using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Draws a single <see cref="PageLayout"/> onto an <see cref="SKCanvas"/>. All spreadsheet-content
    /// drawing happens inside a translated+scaled coordinate group so fit-to-page scale applies
    /// uniformly to cell geometry, borders, and text.
    /// </summary>
    internal sealed class SheetPdfRenderer
    {
        private static readonly SKColor GridlineColor = new SKColor(0xD0, 0xD0, 0xD0);

        private readonly RenderContext _context;
        private readonly Dictionary<ChartModel, ParsedChart> _chartCache = new Dictionary<ChartModel, ParsedChart>();
        private readonly Dictionary<ShapeModel, ParsedShape> _shapeCache = new Dictionary<ShapeModel, ParsedShape>();
        private readonly Dictionary<SmartArtModel, List<SmartArtShape>> _smartArtCache = new Dictionary<SmartArtModel, List<SmartArtShape>>();
        private readonly Dictionary<WorksheetModel, ConditionalFormatEvaluator> _cfCache = new Dictionary<WorksheetModel, ConditionalFormatEvaluator>();
        private readonly RichTextCellRenderer _richTextRenderer;
        private readonly TableStyleResolver _tableStyleResolver;

        public SheetPdfRenderer(RenderContext context)
        {
            _context = context;
            _richTextRenderer = new RichTextCellRenderer(context);
            _tableStyleResolver = new TableStyleResolver();
        }

        public void RenderPage(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;

            canvas.Save();
            canvas.Translate((float)page.ContentOriginXPt, (float)page.ContentOriginYPt);
            canvas.Scale((float)page.ScaleFactor);

            // Clip to the printable region (in content coordinates) so scaled content never bleeds
            // into the page margins.
            var contentWidth = (float)(layout.ColumnStartPt[page.EndColumn + 1] - layout.ColumnStartPt[page.StartColumn]);
            var contentHeight = (float)(layout.RowStartPt[page.EndRow + 1] - layout.RowStartPt[page.StartRow]);
            var contentClip = new SKRect(0f, 0f, contentWidth, contentHeight);

            // Cell grid: clip tightly to the content span.
            canvas.Save();
            canvas.ClipRect(contentClip);

            DrawFills(canvas, page);

            if (layout.Sheet.PageSetup.PrintOptions.GridLines)
            {
                DrawGridlines(canvas, page);
            }

            DrawCellBordersAndText(canvas, page);

            DrawConditionalIcons(canvas, page);

            // Merged regions render last, as self-contained units, so they draw correctly on every
            // page they intersect - including continuation pages where the merge's origin lies on an
            // earlier page. The page-level clip trims the parts that fall outside this page.
            DrawMerges(canvas, page);

            canvas.Restore();

            // Floating objects (pictures, charts, shapes) may slightly overrun the used range (line
            // width, arrowheads) at the sheet's outer edges, where empty page space follows. Allow a
            // small overflow there, but keep the clip tight at internal page-split boundaries so a
            // wide object still splits cleanly across pages.
            // 3D chart series-axis labels can legitimately extend farther beyond the used range than
            // pictures or arrowheads, especially on the sheet's outer right edge. Keep the clip tight
            // at internal page breaks, but allow a larger outer-edge overflow so those labels are not
            // truncated to just their first few characters.
            const float overflow = 72f;
            var objectClip = new SKRect(
                0f,
                0f,
                contentWidth + (page.EndColumn >= layout.LastColumn ? overflow : 0f),
                contentHeight + (page.EndRow >= layout.LastRow ? overflow : 0f));

            canvas.Save();
            canvas.ClipRect(objectClip);
            DrawPictures(canvas, page);
            DrawCharts(canvas, page);
            DrawShapes(canvas, page);
            DrawSmartArt(canvas, page);
            canvas.Restore();

            canvas.Restore();
        }

        private double ContentX(SheetLayout layout, int startCol, int col)
        {
            return layout.ColumnStartPt[col] - layout.ColumnStartPt[startCol];
        }

        private double ContentY(SheetLayout layout, int startRow, int row)
        {
            return layout.RowStartPt[row] - layout.RowStartPt[startRow];
        }

        private ConditionalFormatEvaluator ConditionalFormatFor(SheetLayout layout)
        {
            ConditionalFormatEvaluator evaluator;
            if (!_cfCache.TryGetValue(layout.Sheet, out evaluator))
            {
                evaluator = new ConditionalFormatEvaluator(layout.Sheet, _context);
                _cfCache[layout.Sheet] = evaluator;
            }

            return evaluator;
        }

        private void DrawFills(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;
            var cf = ConditionalFormatFor(layout);
            using (var paint = new SKPaint())
            {
                paint.Style = SKPaintStyle.Fill;
                paint.IsAntialias = false;

                for (var row = page.StartRow; row <= page.EndRow; row++)
                {
                    if (layout.RowHidden[row]) continue;

                    for (var col = page.StartColumn; col <= page.EndColumn; col++)
                    {
                        if (layout.ColumnHidden[col]) continue;
                        if (IsPartOfMerge(layout, row, col)) continue;

                        CellRecord record;
                        layout.Sheet.Cells.TryGetValue(new CellAddress(row, col), out record);
                        var baseStyle = record != null && record.Style != null ? record.Style : StyleValue.Default;
                        var style = _tableStyleResolver.Resolve(layout.Sheet, row, col, baseStyle);

                        // A matching conditional-format fill (dxf or color scale) overrides the cell fill.
                        ConditionalFormatEvaluator.CellEffect effect;
                        if (cf.TryGet(row, col, out effect) && effect.HasFill)
                        {
                            var rect = CellRect(page, row, col);
                            DrawRotatedNeighborUnderlay(canvas, page, row, col, style);
                            DrawSolidCellFill(canvas, rect, style.Alignment, effect.Fill);
                        }
                        else if (style != null && style.Pattern != FillPatternKind.None)
                        {
                            DrawRotatedNeighborUnderlay(canvas, page, row, col, style);
                            DrawPatternCellFill(canvas, CellRect(page, row, col), style);
                        }
                    }
                }
            }

            DrawConditionalDataBars(canvas, page, cf);
        }

        /// <summary>
        /// Draws conditional-formatting data bars over the cell fills and behind the cell text. Each bar
        /// grows from the cell's left edge with a gradient toward a lighter tint, as Excel renders the
        /// default gradient data bar.
        /// </summary>
        private void DrawConditionalDataBars(SKCanvas canvas, PageLayout page, ConditionalFormatEvaluator cf)
        {
            var layout = page.Sheet;
            for (var row = page.StartRow; row <= page.EndRow; row++)
            {
                if (layout.RowHidden[row]) continue;

                for (var col = page.StartColumn; col <= page.EndColumn; col++)
                {
                    if (layout.ColumnHidden[col]) continue;
                    if (IsPartOfMerge(layout, row, col)) continue;

                    ConditionalFormatEvaluator.CellEffect effect;
                    if (!cf.TryGet(row, col, out effect) || effect.Bar == null)
                    {
                        continue;
                    }

                    var bar = effect.Bar.Value;
                    var rect = CellRect(page, row, col);
                    var inset = 1.5f;
                    var maxWidth = rect.Width - 2f * inset;
                    var width = (float)(maxWidth * bar.Fraction);
                    if (width <= 0.5f)
                    {
                        continue;
                    }

                    var barRect = new SKRect(rect.Left + inset, rect.Top + inset, rect.Left + inset + width, rect.Bottom - inset);
                    var lighter = new SKColor(
                        (byte)(bar.Color.Red + (255 - bar.Color.Red) * 0.72f),
                        (byte)(bar.Color.Green + (255 - bar.Color.Green) * 0.72f),
                        (byte)(bar.Color.Blue + (255 - bar.Color.Blue) * 0.72f));
                    using (var shader = SKShader.CreateLinearGradient(
                        new SKPoint(barRect.Left, 0f), new SKPoint(barRect.Right, 0f),
                        new[] { bar.Color, lighter }, null, SKShaderTileMode.Clamp))
                    using (var paint = new SKPaint { Style = SKPaintStyle.Fill, IsAntialias = true, Shader = shader })
                    {
                        canvas.DrawRect(barRect, paint);
                    }

                    using (var border = new SKPaint { Style = SKPaintStyle.Stroke, StrokeWidth = 0.6f, IsAntialias = true, Color = bar.Color })
                    {
                        canvas.DrawRect(barRect, border);
                    }
                }
            }
        }

        /// <summary>
        /// Draws conditional-formatting icon-set glyphs at the left of each cell. Icon-set cells are
        /// numeric and right-aligned, so the left-anchored icon does not collide with the value.
        /// </summary>
        private void DrawConditionalIcons(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;
            var cf = ConditionalFormatFor(layout);
            for (var row = page.StartRow; row <= page.EndRow; row++)
            {
                if (layout.RowHidden[row]) continue;

                for (var col = page.StartColumn; col <= page.EndColumn; col++)
                {
                    if (layout.ColumnHidden[col]) continue;
                    if (IsPartOfMerge(layout, row, col)) continue;

                    ConditionalFormatEvaluator.CellEffect effect;
                    if (!cf.TryGet(row, col, out effect) || effect.IconGlyph == null)
                    {
                        continue;
                    }

                    var rect = CellRect(page, row, col);
                    DrawIconGlyph(canvas, rect, effect.IconGlyph.Value);
                }
            }
        }

        private static void DrawIconGlyph(SKCanvas canvas, SKRect cell, ConditionalFormatEvaluator.Icon icon)
        {
            // Fit the glyph to the row height with a little padding, capped at Excel's ~11pt icon size.
            var size = Math.Min(cell.Height - 2f, 11f);
            if (size <= 2f)
            {
                return;
            }

            var cx = cell.Left + 2f + size / 2f;
            var cy = cell.MidY;
            var half = size / 2f;

            using (var paint = new SKPaint { IsAntialias = true, Color = icon.Color })
            {
                if (!icon.IsArrow)
                {
                    paint.Style = SKPaintStyle.Fill;
                    canvas.DrawCircle(cx, cy, half, paint);
                    return;
                }

                // Build an upward arrow (shaft + head) centred on the origin, then rotate it to the
                // glyph's direction. Canvas y is down, so a positive "up" points to -y.
                canvas.Save();
                canvas.Translate(cx, cy);
                canvas.RotateDegrees(-icon.AngleDeg + 90f); // 90 (up) => 0 rotation
                // Excel's arrow glyphs are chunky (roughly 9 wide x 11 tall for an 11pt icon).
                var shaftW = size * 0.34f;
                var headW = size * 0.82f;
                var headLen = size * 0.5f;
                paint.Style = SKPaintStyle.Fill;
                using (var path = new SKPath())
                {
                    path.MoveTo(0f, -half);                 // tip (up)
                    path.LineTo(headW / 2f, -half + headLen);
                    path.LineTo(shaftW / 2f, -half + headLen);
                    path.LineTo(shaftW / 2f, half);
                    path.LineTo(-shaftW / 2f, half);
                    path.LineTo(-shaftW / 2f, -half + headLen);
                    path.LineTo(-headW / 2f, -half + headLen);
                    path.Close();
                    canvas.DrawPath(path, paint);
                }

                canvas.Restore();
            }
        }

        private void DrawGridlines(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;
            using (var paint = new SKPaint())
            {
                paint.Color = GridlineColor;
                paint.Style = SKPaintStyle.Stroke;
                paint.StrokeWidth = 0.5f;
                paint.IsAntialias = false;

                var top = (float)ContentY(layout, page.StartRow, page.StartRow);
                var bottom = (float)ContentY(layout, page.StartRow, page.EndRow + 1);
                for (var col = page.StartColumn; col <= page.EndColumn + 1; col++)
                {
                    var x = (float)ContentX(layout, page.StartColumn, col);
                    canvas.DrawLine(x, top, x, bottom, paint);
                }

                var left = (float)ContentX(layout, page.StartColumn, page.StartColumn);
                var right = (float)ContentX(layout, page.StartColumn, page.EndColumn + 1);
                for (var row = page.StartRow; row <= page.EndRow + 1; row++)
                {
                    var y = (float)ContentY(layout, page.StartRow, row);
                    canvas.DrawLine(left, y, right, y, paint);
                }
            }
        }

        private void DrawCellBordersAndText(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;

            for (var row = page.StartRow; row <= page.EndRow; row++)
            {
                if (layout.RowHidden[row]) continue;

                for (var col = page.StartColumn; col <= page.EndColumn; col++)
                {
                    if (layout.ColumnHidden[col]) continue;
                    if (IsPartOfMerge(layout, row, col)) continue;

                    CellRecord record;
                    layout.Sheet.Cells.TryGetValue(new CellAddress(row, col), out record);
                    if (record == null)
                    {
                        var emptyStyle = _tableStyleResolver.Resolve(layout.Sheet, row, col, StyleValue.Default);
                        if (emptyStyle == null || !HasVisibleBorder(emptyStyle.Borders))
                        {
                            continue;
                        }

                        DrawBorders(canvas, CellRect(page, row, col), emptyStyle.Borders, BorderShearFactor(emptyStyle.Alignment));
                        continue;
                    }

                    var style = _tableStyleResolver.Resolve(layout.Sheet, row, col, record.Style ?? StyleValue.Default);
                    var rect = CellRect(page, row, col);

                    DrawBorders(canvas, rect, style.Borders, BorderShearFactor(style.Alignment));
                    DrawText(canvas, page, row, col, rect, record, style);
                }
            }
        }

        private SKRect CellRect(PageLayout page, int row, int col)
        {
            var layout = page.Sheet;

            var lastRow = row;
            var lastCol = col;
            MergeRegion merge;
            if (layout.TryGetMergeOrigin(row, col, out merge))
            {
                lastRow = Math.Min(page.EndRow, row + Math.Max(1, merge.TotalRows) - 1);
                lastCol = Math.Min(page.EndColumn, col + Math.Max(1, merge.TotalColumns) - 1);
            }

            var left = (float)ContentX(layout, page.StartColumn, col);
            var top = (float)ContentY(layout, page.StartRow, row);
            var right = (float)ContentX(layout, page.StartColumn, lastCol + 1);
            var bottom = (float)ContentY(layout, page.StartRow, lastRow + 1);
            return new SKRect(left, top, right, bottom);
        }

        private static bool IsPartOfMerge(SheetLayout layout, int row, int col)
        {
            MergeRegion merge;
            return layout.IsMergeCovered(row, col) || layout.TryGetMergeOrigin(row, col, out merge);
        }

        /// <summary>
        /// Renders every merged region that intersects this page as a single unit (fill, borders,
        /// then origin-cell text), using the region's full rectangle so a merge whose origin sits on
        /// an earlier page still paints its continuation here. The active page clip trims overflow.
        /// </summary>
        private void DrawMerges(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;

            foreach (var merge in layout.Sheet.MergeRegions)
            {
                if (!MergeIntersectsPage(page, merge))
                {
                    continue;
                }

                CellRecord record;
                layout.Sheet.Cells.TryGetValue(new CellAddress(merge.FirstRow, merge.FirstColumn), out record);
                var style = _tableStyleResolver.Resolve(layout.Sheet, merge.FirstRow, merge.FirstColumn, record != null && record.Style != null ? record.Style : StyleValue.Default);
                var rect = FullMergeRect(page, merge);

                if (style.Pattern != FillPatternKind.None)
                {
                    DrawPatternCellFill(canvas, rect, style);
                }

                DrawBorders(canvas, rect, style.Borders, BorderShearFactor(style.Alignment));

                // Text is drawn only on the page that contains the merge's origin cell. Excel clips
                // merged-cell text at the page break rather than repeating its tail on the
                // continuation page, which otherwise shows only the fill/borders.
                if (record != null && OriginOnPage(page, merge))
                {
                    DrawText(canvas, page, merge.FirstRow, merge.FirstColumn, rect, record, style);
                }
            }
        }

        private static bool OriginOnPage(PageLayout page, MergeRegion merge)
        {
            return merge.FirstRow >= page.StartRow && merge.FirstRow <= page.EndRow
                && merge.FirstColumn >= page.StartColumn && merge.FirstColumn <= page.EndColumn;
        }

        private void DrawSolidCellFill(SKCanvas canvas, SKRect rect, AlignmentValue alignment, SKColor color)
        {
            if (color.Alpha == 0)
            {
                return;
            }

            var shear = FillShearFactor(alignment);
            if (shear == 0f)
            {
                using (var paint = new SKPaint())
                {
                    paint.Style = SKPaintStyle.Fill;
                    paint.IsAntialias = false;
                    paint.Color = color;
                    canvas.DrawRect(rect, paint);
                }

                return;
            }

            canvas.Save();
            using (var path = CreateShearedCellPath(rect, shear))
            using (var paint = new SKPaint())
            {
                paint.Style = SKPaintStyle.Fill;
                paint.IsAntialias = false;
                paint.Color = color;
                canvas.DrawPath(path, paint);
            }

            canvas.Restore();
        }

        private void DrawPatternCellFill(SKCanvas canvas, SKRect rect, StyleValue style)
        {
            var shear = FillShearFactor(style.Alignment);
            if (shear == 0f)
            {
                PatternFillRenderer.Draw(canvas, rect, style, _context.Colors);
                return;
            }

            canvas.Save();
            using (var path = CreateShearedCellPath(rect, shear))
            {
                canvas.ClipPath(path, SKClipOperation.Intersect, true);
                PatternFillRenderer.Draw(canvas, ShearedBounds(rect, shear), style, _context.Colors);
            }

            canvas.Restore();
        }

        private void DrawRotatedNeighborUnderlay(SKCanvas canvas, PageLayout page, int row, int col, StyleValue style)
        {
            if (style == null || style.Alignment == null)
            {
                return;
            }

            var rotation = style.Alignment.TextRotation;
            if (rotation <= 90 || rotation >= 180 || col >= page.Sheet.LastColumn)
            {
                return;
            }

            var nextStyle = ResolveCellStyle(page.Sheet.Sheet, row, col + 1);
            if (nextStyle == null || nextStyle.Pattern != FillPatternKind.Solid)
            {
                return;
            }

            var fill = _context.Colors.Resolve(nextStyle.ForegroundColor, SKColors.Transparent);
            if (fill.Alpha == 0)
            {
                fill = _context.Colors.Resolve(nextStyle.BackgroundColor, SKColors.Transparent);
            }

            DrawSolidCellFill(canvas, CellRect(page, row, col), null, fill);
        }

        private StyleValue ResolveCellStyle(WorksheetModel sheet, int row, int col)
        {
            CellRecord record;
            sheet.Cells.TryGetValue(new CellAddress(row, col), out record);
            var baseStyle = record != null && record.Style != null ? record.Style : StyleValue.Default;
            return _tableStyleResolver.Resolve(sheet, row, col, baseStyle);
        }

        private static SKRect ShearedBounds(SKRect rect, float shear)
        {
            var dxTop = -shear * rect.Height;
            var left = Math.Min(rect.Left, rect.Left + dxTop);
            var right = Math.Max(rect.Right, rect.Right + dxTop);
            return new SKRect(left, rect.Top, right, rect.Bottom);
        }

        private static float FillShearFactor(AlignmentValue alignment)
        {
            if (alignment == null)
            {
                return 0f;
            }

            var rotation = alignment.TextRotation;
            if (rotation <= 90 || rotation >= 180)
            {
                return 0f;
            }

            return BorderShearFactor(alignment);
        }

        private static SKPath CreateShearedCellPath(SKRect rect, float shear)
        {
            var dxTop = -shear * rect.Height;
            var tl = new SKPoint(rect.Left + dxTop, rect.Top);
            var tr = new SKPoint(rect.Right + dxTop, rect.Top);
            var br = new SKPoint(rect.Right, rect.Bottom);
            var bl = new SKPoint(rect.Left, rect.Bottom);
            var path = new SKPath();
            path.MoveTo(tl);
            path.LineTo(tr);
            path.LineTo(br);
            path.LineTo(bl);
            path.Close();
            return path;
        }

        private static bool MergeIntersectsPage(PageLayout page, MergeRegion merge)
        {
            var lastRow = merge.FirstRow + Math.Max(1, merge.TotalRows) - 1;
            var lastCol = merge.FirstColumn + Math.Max(1, merge.TotalColumns) - 1;
            return merge.FirstRow <= page.EndRow && lastRow >= page.StartRow
                && merge.FirstColumn <= page.EndColumn && lastCol >= page.StartColumn;
        }

        /// <summary>
        /// The merged region's full rectangle in this page's content coordinates. Columns/rows that
        /// precede the page's start yield negative offsets (clipped away by the page clip), which is
        /// what lets a continuation page show only the part of the merge that belongs to it.
        /// </summary>
        private SKRect FullMergeRect(PageLayout page, MergeRegion merge)
        {
            var layout = page.Sheet;
            var firstRow = merge.FirstRow;
            var firstCol = merge.FirstColumn;
            var lastRow = Math.Min(layout.LastRow, merge.FirstRow + Math.Max(1, merge.TotalRows) - 1);
            var lastCol = Math.Min(layout.LastColumn, merge.FirstColumn + Math.Max(1, merge.TotalColumns) - 1);

            var left = (float)ContentX(layout, page.StartColumn, firstCol);
            var top = (float)ContentY(layout, page.StartRow, firstRow);
            var right = (float)ContentX(layout, page.StartColumn, lastCol + 1);
            var bottom = (float)ContentY(layout, page.StartRow, lastRow + 1);
            return new SKRect(left, top, right, bottom);
        }

        /// <summary>English Metric Units per point (914400 EMU per inch / 72 points per inch).</summary>
        private const double EmuPerPoint = 12700d;

        /// <summary>
        /// Draws each anchored picture that intersects this page at its EMU-precise rectangle. Images
        /// are decoded on demand; the active page clip trims any part outside the page.
        /// </summary>
        private void DrawPictures(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;

            foreach (var picture in layout.Sheet.Pictures)
            {
                if (picture == null || picture.ImageData == null || picture.ImageData.Length == 0)
                {
                    continue;
                }

                var rect = PictureRect(page, picture);
                if (rect.Width <= 0f || rect.Height <= 0f)
                {
                    continue;
                }

                if (!PictureRectIntersectsPage(page, rect))
                {
                    continue;
                }

                var image = _context.Pictures.GetImage(picture, rect.Width, rect.Height);
                if (image == null)
                {
                    continue;
                }

                using (var paint = new SKPaint())
                {
                    paint.IsAntialias = true;
                    paint.FilterQuality = PictureFilterQuality(picture);
                    canvas.DrawImage(image, rect, paint);
                }
            }
        }

        private static SKFilterQuality PictureFilterQuality(PictureModel picture)
        {
            if (picture == null || string.IsNullOrEmpty(picture.ImageExtension))
            {
                return SKFilterQuality.High;
            }

            var extension = picture.ImageExtension.Trim().ToLowerInvariant();
            if (extension == "png" || extension == "gif" || extension == "bmp")
            {
                // UI screenshots and text-heavy raster artwork stay sharper with minimal resampling.
                return SKFilterQuality.None;
            }

            return SKFilterQuality.High;
        }

        private static bool PictureRectIntersectsPage(PageLayout page, SKRect rect)
        {
            var layout = page.Sheet;
            var contentWidth = (float)(layout.ColumnStartPt[page.EndColumn + 1] - layout.ColumnStartPt[page.StartColumn]);
            var contentHeight = (float)(layout.RowStartPt[page.EndRow + 1] - layout.RowStartPt[page.StartRow]);
            return rect.Right > 0f && rect.Bottom > 0f && rect.Left < contentWidth && rect.Top < contentHeight;
        }

        /// <summary>
        /// Draws each chart intersecting this page at its EMU anchor rectangle. A chart whose XML is
        /// unsupported is skipped. The active page clip trims parts outside this page.
        /// </summary>
        private void DrawCharts(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;
            if (layout.Sheet.Charts.Count == 0)
            {
                return;
            }

            var renderer = new ChartRenderer(_context);
            foreach (var chart in layout.Sheet.Charts)
            {
                if (chart == null || string.IsNullOrEmpty(chart.RawChartXml))
                {
                    continue;
                }

                if (chart.UpperLeftRow > page.EndRow || chart.LowerRightRow < page.StartRow
                    || chart.UpperLeftColumn > page.EndColumn || chart.LowerRightColumn < page.StartColumn)
                {
                    continue;
                }

                ParsedChart parsed;
                if (!_chartCache.TryGetValue(chart, out parsed))
                {
                    parsed = ChartXmlParser.Parse(chart.RawChartXml, _context.Colors, DateSystemOf(), _context.Culture, _context.Workbook, chart);
                    _chartCache[chart] = parsed;
                }

                if (parsed == null)
                {
                    continue;
                }

                var rect = ChartRect(page, chart);
                if (rect.Width > 0f && rect.Height > 0f)
                {
                    // Keep charts vector-backed in the PDF so thin series strokes, gridlines, and
                    // axis labels stay crisp instead of being flattened into a raster snapshot.
                    renderer.Draw(canvas, page, rect, parsed, false);
                }
            }
        }

        /// <summary>
        /// Draws each SmartArt diagram intersecting this page. The diagram's pre-laid-out shapes
        /// (from its dsp:drawing part) are positioned relative to the diagram frame's anchor and
        /// rendered with the shared shape renderer - no SmartArt layout engine required.
        /// </summary>
        private void DrawSmartArt(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;
            if (layout.Sheet.SmartArts.Count == 0)
            {
                return;
            }

            var renderer = new ShapeRenderer(_context);
            foreach (var smartArt in layout.Sheet.SmartArts)
            {
                if (smartArt == null || string.IsNullOrEmpty(smartArt.RawDrawingXml))
                {
                    continue;
                }

                if (smartArt.UpperLeftRow > page.EndRow || smartArt.LowerRightRow < page.StartRow
                    || smartArt.UpperLeftColumn > page.EndColumn || smartArt.LowerRightColumn < page.StartColumn)
                {
                    continue;
                }

                List<SmartArtShape> shapes;
                if (!_smartArtCache.TryGetValue(smartArt, out shapes))
                {
                    shapes = ShapeXmlParser.ParseSmartArtDrawing(smartArt.RawDrawingXml, _context.Colors);
                    _smartArtCache[smartArt] = shapes;
                }

                if (shapes.Count == 0)
                {
                    continue;
                }

                var originX = layout.ColumnStartPt[page.StartColumn];
                var originY = layout.RowStartPt[page.StartRow];

                // Diagram origin: the column offset applies to X, but the dsp shape coordinates
                // already bake in the row offset (their vertical origin is the anchor cell's top),
                // so Y uses the cell top without the row offset.
                var diagramLeft = ColumnEdgePt(layout, smartArt.UpperLeftColumn, smartArt.UpperLeftColumnOffset) - originX;
                var diagramTop = layout.RowStartPt[smartArt.UpperLeftRow] - originY;
                var frameRight = ColumnEdgePt(layout, smartArt.LowerRightColumn, smartArt.LowerRightColumnOffset) - originX;
                var frameWidth = frameRight - diagramLeft;

                // Excel scales the laid-out diagram to fit the frame width.
                double minX, minY, maxX, maxY;
                CanvasBounds(shapes, out minX, out minY, out maxX, out maxY);
                if (maxX <= 0d)
                {
                    continue;
                }

                var scale = frameWidth > 0d ? frameWidth / maxX : 1d;

                foreach (var s in shapes)
                {
                    if (s.Shape == null)
                    {
                        continue;
                    }

                    var left = (float)(diagramLeft + s.XPt * scale);
                    var top = (float)(diagramTop + s.YPt * scale);
                    var rect = new SKRect(left, top, left + (float)(s.WPt * scale), top + (float)(s.HPt * scale));
                    if (rect.Width > 0f && rect.Height > 0f)
                    {
                        renderer.Draw(canvas, page, rect, s.Shape, null, null);
                    }
                }
            }
        }

        /// <summary>
        /// The bounding box of the laid-out SmartArt shapes, accounting for each shape's rotation
        /// (a 90-degree-rotated shape contributes its swapped width/height).
        /// </summary>
        private static void CanvasBounds(List<SmartArtShape> shapes, out double minX, out double minY, out double maxX, out double maxY)
        {
            minX = double.MaxValue;
            minY = double.MaxValue;
            maxX = double.MinValue;
            maxY = double.MinValue;

            foreach (var s in shapes)
            {
                if (s.Shape == null)
                {
                    continue;
                }

                var cx = s.XPt + s.WPt / 2d;
                var cy = s.YPt + s.HPt / 2d;
                var rad = s.Shape.RotationDeg * Math.PI / 180d;
                var cos = Math.Abs(Math.Cos(rad));
                var sin = Math.Abs(Math.Sin(rad));
                var halfW = (s.WPt * cos + s.HPt * sin) / 2d;
                var halfH = (s.WPt * sin + s.HPt * cos) / 2d;

                if (cx - halfW < minX) minX = cx - halfW;
                if (cy - halfH < minY) minY = cy - halfH;
                if (cx + halfW > maxX) maxX = cx + halfW;
                if (cy + halfH > maxY) maxY = cy + halfH;
            }
        }

        private Aspose.Cells_FOSS.Core.DateSystem DateSystemOf()
        {
            return _context.Workbook != null && _context.Workbook.Settings != null
                ? _context.Workbook.Settings.DateSystem
                : Aspose.Cells_FOSS.Core.DateSystem.Windows1900;
        }

        private SKRect ChartRect(PageLayout page, ChartModel chart)
        {
            var layout = page.Sheet;
            var originX = layout.ColumnStartPt[page.StartColumn];
            var originY = layout.RowStartPt[page.StartRow];

            var left = ColumnEdgePt(layout, chart.UpperLeftColumn, chart.UpperLeftColumnOffset) - originX;
            var top = RowEdgePt(layout, chart.UpperLeftRow, chart.UpperLeftRowOffset) - originY;
            var right = ColumnEdgePt(layout, chart.LowerRightColumn, chart.LowerRightColumnOffset) - originX;
            var bottom = RowEdgePt(layout, chart.LowerRightRow, chart.LowerRightRowOffset) - originY;
            return new SKRect((float)left, (float)top, (float)right, (float)bottom);
        }

        /// <summary>
        /// Draws each drawing shape (autoShape or connector) intersecting this page at its anchor
        /// rectangle. Shapes whose geometry is unknown fall back to a rectangle.
        /// </summary>
        private void DrawShapes(SKCanvas canvas, PageLayout page)
        {
            var layout = page.Sheet;
            if (layout.Sheet.Shapes.Count == 0)
            {
                return;
            }

            var renderer = new ShapeRenderer(_context);
            foreach (var shape in layout.Sheet.Shapes)
            {
                if (shape == null)
                {
                    continue;
                }

                if (shape.UpperLeftRow > page.EndRow || shape.LowerRightRow < page.StartRow
                    || shape.UpperLeftColumn > page.EndColumn || shape.LowerRightColumn < page.StartColumn)
                {
                    continue;
                }

                ParsedShape parsed;
                if (!_shapeCache.TryGetValue(shape, out parsed))
                {
                    parsed = ShapeXmlParser.Parse(shape.RawElementXml, shape.GeometryType, _context.Colors);
                    _shapeCache[shape] = parsed;
                }

                if (parsed == null)
                {
                    continue;
                }

                var layoutRef = page.Sheet;
                var originX = layoutRef.ColumnStartPt[page.StartColumn];
                var originY = layoutRef.RowStartPt[page.StartRow];
                var left = ColumnEdgePt(layoutRef, shape.UpperLeftColumn, shape.UpperLeftColumnOffset) - originX;
                var top = RowEdgePt(layoutRef, shape.UpperLeftRow, shape.UpperLeftRowOffset) - originY;
                var right = ColumnEdgePt(layoutRef, shape.LowerRightColumn, shape.LowerRightColumnOffset) - originX;
                var bottom = RowEdgePt(layoutRef, shape.LowerRightRow, shape.LowerRightRowOffset) - originY;
                var rect = new SKRect((float)left, (float)top, (float)right, (float)bottom);
                if (rect.Width > 0f && rect.Height > 0f)
                {
                    renderer.Draw(canvas, page, rect, parsed, shape, layout.Sheet.Shapes);
                }
            }
        }

        private SKRect PictureRect(PageLayout page, PictureModel picture)
        {
            var layout = page.Sheet;
            var originX = layout.ColumnStartPt[page.StartColumn];
            var originY = layout.RowStartPt[page.StartRow];

            var left = ColumnEdgePt(layout, picture.UpperLeftColumn, picture.UpperLeftColumnOffset) - originX;
            var top = RowEdgePt(layout, picture.UpperLeftRow, picture.UpperLeftRowOffset) - originY;
            double right;
            double bottom;
            if (picture.ExtentCx > 0)
            {
                right = left + picture.ExtentCx / EmuPerPoint;
            }
            else
            {
                right = ColumnEdgePt(layout, picture.LowerRightColumn, picture.LowerRightColumnOffset) - originX;
            }

            if (picture.ExtentCy > 0)
            {
                bottom = top + picture.ExtentCy / EmuPerPoint;
            }
            else
            {
                bottom = RowEdgePt(layout, picture.LowerRightRow, picture.LowerRightRowOffset) - originY;
            }

            return new SKRect((float)left, (float)top, (float)right, (float)bottom);
        }

        private static double ColumnEdgePt(SheetLayout layout, int column, long emuOffset)
        {
            if (column < 0) column = 0;
            if (column > layout.LastColumn + 1) column = layout.LastColumn + 1;
            return layout.ColumnStartPt[column] + emuOffset / EmuPerPoint;
        }

        private static double RowEdgePt(SheetLayout layout, int row, long emuOffset)
        {
            if (row < 0) row = 0;
            if (row > layout.LastRow + 1) row = layout.LastRow + 1;
            return layout.RowStartPt[row] + emuOffset / EmuPerPoint;
        }

        /// <summary>
        /// Extends a cell's text clip rectangle across consecutive empty neighbor columns (within the
        /// page) so long non-wrapped text spills over the way Excel renders it. Left/general text
        /// extends right, right-aligned extends left, centered extends both ways.
        /// </summary>
        private SKRect ExpandClipForOverflow(PageLayout page, int row, int col, SKRect rect, HorizontalAlignment horizontal)
        {
            var layout = page.Sheet;
            var left = rect.Left;
            var right = rect.Right;

            var extendRight = horizontal == HorizontalAlignment.Left
                || horizontal == HorizontalAlignment.General
                || horizontal == HorizontalAlignment.Fill
                || horizontal == HorizontalAlignment.Justify
                || horizontal == HorizontalAlignment.Center
                || horizontal == HorizontalAlignment.CenterContinuous;

            var extendLeft = horizontal == HorizontalAlignment.Right
                || horizontal == HorizontalAlignment.Center
                || horizontal == HorizontalAlignment.CenterContinuous;

            if (extendRight)
            {
                var c = col + 1;
                while (c <= page.EndColumn && IsCellEmptyForOverflow(layout, row, c))
                {
                    right = (float)ContentX(layout, page.StartColumn, c + 1);
                    c++;
                }
            }

            if (extendLeft)
            {
                var c = col - 1;
                while (c >= page.StartColumn && IsCellEmptyForOverflow(layout, row, c))
                {
                    left = (float)ContentX(layout, page.StartColumn, c);
                    c--;
                }
            }

            return new SKRect(left, rect.Top, right, rect.Bottom);
        }

        private bool IsCellEmptyForOverflow(SheetLayout layout, int row, int col)
        {
            if (layout.ColumnHidden[col] || layout.IsMergeCovered(row, col))
            {
                return false;
            }

            MergeRegion merge;
            if (layout.TryGetMergeOrigin(row, col, out merge))
            {
                return false;
            }

            CellRecord record;
            if (!layout.Sheet.Cells.TryGetValue(new CellAddress(row, col), out record) || record == null)
            {
                return true;
            }

            var style = record.Style ?? StyleValue.Default;
            var text = DisplayTextFormatter.FormatDisplayValue(record.Value, style, _context.Culture);
            return string.IsNullOrEmpty(text);
        }

        private void DrawBorders(SKCanvas canvas, SKRect rect, BordersValue borders)
        {
            DrawBorders(canvas, rect, borders, 0f);
        }

        /// <summary>
        /// Draws the four cell borders. When <paramref name="shear"/> is non-zero (a cell with rotated
        /// text), Excel skews the box into a parallelogram whose side edges run parallel to the text
        /// baseline: the top/bottom edges stay horizontal and the top edge slides horizontally while the
        /// bottom edge is the pivot (x' = x - shear*height at the top, 0 at the bottom).
        /// </summary>
        private void DrawBorders(SKCanvas canvas, SKRect rect, BordersValue borders, float shear)
        {
            var dxTop = -shear * rect.Height;
            var tlx = rect.Left + dxTop;
            var trx = rect.Right + dxTop;

            DrawBorderSide(canvas, borders.Top, tlx, rect.Top, trx, rect.Top);
            DrawBorderSide(canvas, borders.Bottom, rect.Left, rect.Bottom, rect.Right, rect.Bottom);
            DrawBorderSide(canvas, borders.Left, tlx, rect.Top, rect.Left, rect.Bottom);
            DrawBorderSide(canvas, borders.Right, trx, rect.Top, rect.Right, rect.Bottom);
            if (borders.DiagonalUp)
            {
                DrawBorderSide(canvas, borders.Diagonal, rect.Left, rect.Bottom, trx, rect.Top);
            }

            if (borders.DiagonalDown)
            {
                DrawBorderSide(canvas, borders.Diagonal, tlx, rect.Top, rect.Right, rect.Bottom);
            }
        }

        /// <summary>
        /// Excel's border shear factor for a cell with rotated text: side edges become parallel to the
        /// text baseline (k = -tan(alpha), alpha = signed counter-clockwise angle). Straight and stacked
        /// text leave the box rectangular.
        /// </summary>
        private static float BorderShearFactor(AlignmentValue alignment)
        {
            var rotation = alignment.TextRotation;
            if (rotation <= 0 || rotation >= 255)
            {
                return 0f;
            }

            var ccwDeg = rotation <= 90 ? rotation : -(rotation - 90);
            return (float)(-Math.Tan(ccwDeg * Math.PI / 180d));
        }

        private static bool HasVisibleBorder(BordersValue borders)
        {
            return borders != null
                && ((borders.Top != null && borders.Top.Style != BorderStyle.None)
                    || (borders.Bottom != null && borders.Bottom.Style != BorderStyle.None)
                    || (borders.Left != null && borders.Left.Style != BorderStyle.None)
                    || (borders.Right != null && borders.Right.Style != BorderStyle.None)
                    || ((borders.DiagonalUp || borders.DiagonalDown) && borders.Diagonal != null && borders.Diagonal.Style != BorderStyle.None));
        }

        private void DrawBorderSide(SKCanvas canvas, BorderSideValue side, float x0, float y0, float x1, float y1)
        {
            if (side == null || side.Style == BorderStyle.None)
            {
                return;
            }

            var width = BorderWidthPt(side.Style);
            var color = _context.Colors.Resolve(side.Color, SKColors.Black);

            using (var paint = new SKPaint())
            {
                paint.Color = color;
                paint.Style = SKPaintStyle.Stroke;
                paint.StrokeWidth = width;
                paint.IsAntialias = false;
                paint.StrokeCap = SKStrokeCap.Butt;

                if (side.Style == BorderStyle.SlantedDashDot && x0 != x1 && y0 != y1)
                {
                    paint.IsAntialias = true;
                    paint.StrokeCap = SKStrokeCap.Round;
                }

                var dash = DashFor(side.Style, width);
                if (dash != null)
                {
                    paint.PathEffect = dash;
                }

                if (side.Style == BorderStyle.Double)
                {
                    // Two parallel hairlines with a gap approximates a double border.
                    var offset = width;
                    if (y0 == y1)
                    {
                        canvas.DrawLine(x0, y0 - offset, x1, y1 - offset, paint);
                        canvas.DrawLine(x0, y0 + offset, x1, y1 + offset, paint);
                    }
                    else
                    {
                        canvas.DrawLine(x0 - offset, y0, x1 - offset, y1, paint);
                        canvas.DrawLine(x0 + offset, y0, x1 + offset, y1, paint);
                    }
                }
                else
                {
                    canvas.DrawLine(x0, y0, x1, y1, paint);
                }

                if (dash != null)
                {
                    dash.Dispose();
                    paint.PathEffect = null;
                }
            }
        }

        private static float BorderWidthPt(BorderStyle style)
        {
            switch (style)
            {
                case BorderStyle.Hair:
                    return 0.5f;
                case BorderStyle.Thin:
                case BorderStyle.Dotted:
                case BorderStyle.Dashed:
                case BorderStyle.DashDot:
                case BorderStyle.DashDotDot:
                    return 0.75f;
                case BorderStyle.Medium:
                case BorderStyle.MediumDashed:
                case BorderStyle.MediumDashDot:
                case BorderStyle.MediumDashDotDot:
                case BorderStyle.SlantedDashDot:
                    return 1.75f;
                case BorderStyle.Thick:
                    return 2.5f;
                case BorderStyle.Double:
                    return 0.75f;
                default:
                    return 0.75f;
            }
        }

        private static SKPathEffect DashFor(BorderStyle style, float width)
        {
            switch (style)
            {
                case BorderStyle.Dotted:
                    return SKPathEffect.CreateDash(new[] { 0.9f * width, 1.6f * width }, 0f);
                case BorderStyle.Dashed:
                case BorderStyle.MediumDashed:
                    return SKPathEffect.CreateDash(new[] { 4.8f * width, 2.2f * width }, 0f);
                case BorderStyle.DashDot:
                case BorderStyle.MediumDashDot:
                    return SKPathEffect.CreateDash(new[] { 5.4f * width, 2.3f * width, 1.1f * width, 2.3f * width }, 0f);
                case BorderStyle.DashDotDot:
                case BorderStyle.MediumDashDotDot:
                    return SKPathEffect.CreateDash(new[] { 5.2f * width, 2.1f * width, width, 1.9f * width, width, 2.1f * width }, 0f);
                case BorderStyle.SlantedDashDot:
                    return SKPathEffect.CreateDash(new[] { 4.1f * width, 1.55f * width, 0.95f * width, 1.55f * width }, 0f);
                default:
                    return null;
            }
        }

        private void DrawText(SKCanvas canvas, PageLayout page, int row, int col, SKRect rect, CellRecord record, StyleValue style)
        {
            var fontContext = _context.GetFontContext(style.Font);
            var padding = (float)SheetLayout.HorizontalPaddingPt;
            var baseInnerWidth = rect.Width - 2f * padding;
            if (baseInnerWidth <= 0f)
            {
                return;
            }

            var horizontal = ResolveHorizontal(style.Alignment.Horizontal, record);
            var indentOffset = IndentOffsetPt(fontContext, style.Alignment, horizontal);
            var innerWidth = Math.Max(0f, baseInnerWidth - indentOffset);
            var wrapWidth = EffectiveWrapWidthPt(innerWidth, style.Alignment);
            if (innerWidth <= 0f)
            {
                return;
            }

            // An icon-set rule with showValue="0" replaces the value with just the icon.
            ConditionalFormatEvaluator.CellEffect cf;
            var hasCf = ConditionalFormatFor(page.Sheet).TryGet(row, col, out cf);
            if (hasCf && cf.SuppressText)
            {
                return;
            }

            // Excel's General number format is width-aware: it rounds a number to as many digits as
            // fit the cell (capped at 11 significant digits) rather than showing full precision and
            // clipping. Explicit number formats keep their fixed display and clip if too wide.
            string text;
            if (!style.Alignment.WrapText && IsGeneralNumeric(record.Value, style))
            {
                text = FitGeneralNumber(record.Value, innerWidth, fontContext);
            }
            else
            {
                text = DisplayTextFormatter.FormatDisplayValue(record.Value, style, _context.Culture);
            }

            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            // Rotated / vertically-stacked cell text takes a dedicated path.
            var rotation = style.Alignment.TextRotation;
            if (rotation == 255)
            {
                DrawStackedCellText(canvas, rect, StripLineBreaks(text), fontContext, style);
                return;
            }
            if (rotation != 0)
            {
                DrawRotatedCellText(canvas, rect, StripLineBreaks(text), fontContext, style, rotation);
                return;
            }

            // Wrapped text breaks on explicit newlines; non-wrapped text collapses them the way
            // Excel does (a manual line break renders as zero-width, keeping one visual line)
            // rather than drawing a missing-glyph box for the control character.
            var lines = style.Alignment.WrapText
                ? fontContext.WrapLines(text, wrapWidth)
                : new List<string>(new[] { StripLineBreaks(text) });

            var lineHeight = fontContext.LineHeightPt;
            var blockHeight = lines.Count * lineHeight;

            float blockTop;
            switch (style.Alignment.Vertical)
            {
                case VerticalAlignment.Top:
                    blockTop = rect.Top + (float)SheetLayout.VerticalPaddingPt;
                    break;
                case VerticalAlignment.Center:
                    blockTop = rect.Top + (rect.Height - blockHeight) / 2f;
                    break;
                default: // Bottom / Justify / Distributed
                    blockTop = rect.Bottom - (float)SheetLayout.VerticalPaddingPt - blockHeight;
                    break;
            }

            var color = _context.Colors.Resolve(style.Font.Color, SKColors.Black);
            if (hasCf && cf.HasFontColor)
            {
                color = cf.FontColor;
            }
            // Non-wrapped text that is wider than its cell overflows into adjacent EMPTY cells, the
            // way Excel renders it; wrapped text and text with occupied neighbors stay clipped.
            var clip = rect;
            if (!style.Alignment.WrapText && lines.Count == 1)
            {
                var needed = fontContext.Measure(lines[0]) + 2f * padding + indentOffset;
                if (needed > rect.Width)
                {
                    clip = ExpandClipForOverflow(page, row, col, rect, horizontal);
                }
            }

            if (_richTextRenderer.CanRender(record, style, text))
            {
                if (_richTextRenderer.TryDraw(canvas, page, rect, clip, record, style, horizontal, hasCf && cf.HasFontColor, color))
                {
                    return;
                }
            }

            canvas.Save();
            canvas.ClipRect(clip);

            using (var paint = new SKPaint())
            {
                paint.Color = color;
                paint.IsAntialias = true;

                var metricsFont = new SKFont(_context.Fonts.Resolve(style.Font), fontContext.SizePt);
                var metrics = metricsFont.Metrics;
                metricsFont.Dispose();

                for (var i = 0; i < lines.Count; i++)
                {
                    var line = lines[i];
                    var lineTop = blockTop + i * lineHeight;
                    var baseline = lineTop - metrics.Ascent; // Ascent is negative in Skia.

                    var lineWidth = fontContext.Measure(line);
                    var lineScale = ShrinkScale(style.Alignment, lineWidth, innerWidth);
                    var renderedWidth = lineWidth * lineScale;
                    float x;
                    switch (horizontal)
                    {
                        case HorizontalAlignment.Center:
                        case HorizontalAlignment.CenterContinuous:
                            x = rect.Left + (rect.Width - renderedWidth) / 2f;
                            break;
                        case HorizontalAlignment.Right:
                            x = rect.Right - padding - indentOffset - renderedWidth;
                            break;
                        default:
                            x = rect.Left + padding + indentOffset;
                            break;
                    }

                    if (lineScale < 0.999f)
                    {
                        canvas.Save();
                        canvas.Translate(x, baseline);
                        canvas.Scale(lineScale, lineScale);
                        DrawRuns(canvas, fontContext, line, 0f, 0f, paint, style.Font);
                        canvas.Restore();
                    }
                    else if (!TryRecordRuns(page, clip, fontContext, line, x, baseline, paint.Color))
                    {
                        DrawRuns(canvas, fontContext, line, x, baseline, paint, style.Font);
                    }
                    else
                    {
                        DrawDecorations(canvas, fontContext, line, x, baseline, paint, style.Font);
                    }
                }
            }

            canvas.Restore();
        }

        private static float EffectiveWrapWidthPt(float availableWidth, AlignmentValue alignment)
        {
            if (availableWidth <= 0f || alignment == null)
            {
                return availableWidth;
            }

            if (alignment.Horizontal == HorizontalAlignment.Distributed)
            {
                var indentUnits = alignment.IndentLevel + Math.Max(0, alignment.RelativeIndent);
                if (indentUnits > 0)
                {
                    return availableWidth / (1f + 2f * indentUnits);
                }
            }

            return availableWidth;
        }

        private static float ShrinkScale(AlignmentValue alignment, float textWidth, float availableWidth)
        {
            if (alignment == null || !alignment.ShrinkToFit)
            {
                return 1f;
            }

            if (textWidth <= 0f || availableWidth <= 0f)
            {
                return 1f;
            }

            if (textWidth <= availableWidth)
            {
                return 1f;
            }

            var scale = availableWidth / textWidth;
            if (scale < 0.1f)
            {
                return 0.1f;
            }

            return scale;
        }

        private static float IndentOffsetPt(CellFontContext fontContext, AlignmentValue alignment, HorizontalAlignment horizontal)
        {
            if (fontContext == null || alignment == null)
            {
                return 0f;
            }

            if (alignment.IndentLevel <= 0)
            {
                return 0f;
            }

            if (horizontal != HorizontalAlignment.Left
                && horizontal != HorizontalAlignment.Right
                && horizontal != HorizontalAlignment.Distributed
                && horizontal != HorizontalAlignment.Justify)
            {
                return 0f;
            }

            var digitWidth = fontContext.Measure("0");
            var indentStep = Math.Max(fontContext.SizePt * 0.68f, digitWidth * 1.15f);
            return alignment.IndentLevel * indentStep;
        }

        private bool TryRecordRuns(PageLayout page, SKRect clip, CellFontContext fontContext, string line, float x, float baseline, SKColor color)
        {
            var session = _context.PdfTextSession;
            if (session == null || !_context.EnableWorksheetTextOptimization || string.IsNullOrEmpty(line))
            {
                return false;
            }

            var runs = fontContext.SplitRuns(line);
            for (var i = 0; i < runs.Count; i++)
            {
                if (!SupportsType3Run(runs[i].Typeface))
                {
                    return false;
                }
            }

            var cursor = x;
            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                session.RecordWorksheetRun(page, clip, run.Text, run.Typeface, color, fontContext.SizePt, cursor, baseline);
                cursor += run.WidthPt;
            }

            return true;
        }

        private static bool SupportsType3Run(SKTypeface typeface)
        {
            if (typeface == null || string.IsNullOrEmpty(typeface.FamilyName))
            {
                return false;
            }

            return !string.Equals(typeface.FamilyName, "DengXian", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(typeface.FamilyName, "等线", StringComparison.Ordinal);
        }

        private void DrawRuns(SKCanvas canvas, CellFontContext fontContext, string line, float x, float baseline, SKPaint paint, FontValue font)
        {
            var runs = fontContext.SplitRuns(line);
            var cursor = x;

            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                using (var runFont = new SKFont(run.Typeface, fontContext.SizePt))
                {
                    runFont.Subpixel = true;
                    PdfTextPathRenderer.DrawText(canvas, run.Text, cursor, baseline, runFont, paint.Color);

                    if (font != null && (font.Underline != FontUnderlineType.None))
                    {
                        DrawUnderline(canvas, cursor, baseline, cursor + run.WidthPt, fontContext.SizePt, paint.Color);
                    }

                    if (font != null && font.StrikeThrough)
                    {
                        DrawStrikeThrough(canvas, cursor, baseline, cursor + run.WidthPt, fontContext.SizePt, paint.Color);
                    }
                }

                cursor += run.WidthPt;
            }
        }

        private void DrawDecorations(SKCanvas canvas, CellFontContext fontContext, string line, float x, float baseline, SKPaint paint, FontValue font)
        {
            if (font == null)
            {
                return;
            }

            if (font.Underline == FontUnderlineType.None && !font.StrikeThrough)
            {
                return;
            }

            var runs = fontContext.SplitRuns(line);
            var cursor = x;
            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                if (font.Underline != FontUnderlineType.None)
                {
                    DrawUnderline(canvas, cursor, baseline, cursor + run.WidthPt, fontContext.SizePt, paint.Color);
                }

                if (font.StrikeThrough)
                {
                    DrawStrikeThrough(canvas, cursor, baseline, cursor + run.WidthPt, fontContext.SizePt, paint.Color);
                }

                cursor += run.WidthPt;
            }
        }

        private static void DrawUnderline(SKCanvas canvas, float x0, float baseline, float x1, float fontSizePt, SKColor color)
        {
            using (var paint = new SKPaint())
            {
                paint.Color = color;
                paint.Style = SKPaintStyle.Stroke;
                paint.IsAntialias = true;
                paint.StrokeWidth = Math.Max(0.6f, fontSizePt * 0.055f);
                var y = baseline + fontSizePt * 0.08f;
                canvas.DrawLine(x0, y, x1, y, paint);
            }
        }

        private static void DrawStrikeThrough(SKCanvas canvas, float x0, float baseline, float x1, float fontSizePt, SKColor color)
        {
            using (var paint = new SKPaint())
            {
                paint.Color = color;
                paint.Style = SKPaintStyle.Stroke;
                paint.IsAntialias = true;
                paint.StrokeWidth = Math.Max(0.55f, fontSizePt * 0.05f);
                var y = baseline - fontSizePt * 0.28f;
                canvas.DrawLine(x0, y, x1, y, paint);
            }
        }

        /// <summary>
        /// Draws a single line of cell text rotated by Excel's textRotation value (1-90 =
        /// counter-clockwise degrees above the horizon, 91-180 = clockwise degrees below it). The
        /// rotated text block is aligned within the cell per the cell's horizontal/vertical alignment.
        /// </summary>
        private void DrawRotatedCellText(SKCanvas canvas, SKRect rect, string text, CellFontContext fontContext, StyleValue style, int rotation)
        {
            // Excel angle (counter-clockwise positive); canvas rotation is clockwise-positive (y-down).
            var ccwDeg = rotation <= 90 ? rotation : -(rotation - 90);
            var canvasDeg = (float)(-ccwDeg);
            var isClockwiseDiagonal = rotation > 90 && rotation < 180;

            var color = _context.Colors.Resolve(style.Font.Color, SKColors.Black);
            var padding = (float)SheetLayout.HorizontalPaddingPt;

            using (var paint = new SKPaint { Color = color, IsAntialias = true })
            {
                var metricsFont = new SKFont(_context.Fonts.Resolve(style.Font), fontContext.SizePt);
                var metrics = metricsFont.Metrics;
                metricsFont.Dispose();

                var textW = fontContext.Measure(text);
                var ascent = -metrics.Ascent;
                var descent = metrics.Descent;

                var rad = ccwDeg * Math.PI / 180d;
                var absCos = (float)Math.Abs(Math.Cos(rad));
                var absSin = (float)Math.Abs(Math.Sin(rad));

                // A rotated cell is skewed into a parallelogram only when it actually has a border to
                // draw; a borderless rotated cell keeps its upright rectangle (and its text may spill
                // into neighbours), so the shear-based centring and clip apply only when bordered.
                var shear = HasVisibleBorder(style.Borders) ? BorderShearFactor(style.Alignment) : 0f;

                // The cell can hold text up to this length along the rotation direction; beyond it the
                // text overflows. Excel centers text that fits and left/bottom-anchors text that does not.
                var capacity = Math.Min(
                    absCos > 1e-4f ? (rect.Width - 2f * padding) / absCos : float.MaxValue,
                    absSin > 1e-4f ? rect.Height / absSin : float.MaxValue);
                if (isClockwiseDiagonal && shear != 0f)
                {
                    // Excel is much less willing to center clockwise diagonal text (91-179): the
                    // sheared visible box behaves like a narrower slot, so strings such as "Rot 135"
                    // overflow into the reading-start corner instead of staying fully centered.
                    capacity *= 0.55f;
                }

                float anchorX;
                float anchorY;
                float tx;
                float baseline;
                var allowClockwiseOverflow = isClockwiseDiagonal && shear == 0f;
                if (textW <= capacity && !allowClockwiseOverflow)
                {
                    // Fits: center the block. Horizontally, a bordered (sheared) cell places the text a
                    // fifth of the shear extent left of the upright centre - between that centre and the
                    // parallelogram centroid - so it clears both slanted side borders (matching Excel).
                    // Vertically, Excel centres the visible glyphs rather than the nominal line box, which
                    // sits half the ascender/descender gap below the row centre for ascender-only text.
                    baseline = (ascent - descent) / 2f;
                    anchorX = rect.MidX - shear * rect.Height / 5f;
                    anchorY = rect.MidY + baseline;
                    tx = -textW / 2f;
                }
                else
                {
                    // Overflows: pin the text's rotated bounding box by its reading-start corner so the
                    // start stays visible and the glyphs (whose ascenders lean across the baseline) do not
                    // spill past the cell edge. Clockwise text (down-right) starts at the top edge,
                    // counter-clockwise (up-right) at the bottom.
                    var vpad = (float)SheetLayout.VerticalPaddingPt;
                    var cos = (float)Math.Cos(canvasDeg * Math.PI / 180d);
                    var sin = (float)Math.Sin(canvasDeg * Math.PI / 180d);

                    // Axis-aligned bounds of the layout box (x in [0, textW], y in [-ascent, descent],
                    // baseline at y = 0) after rotation about the origin.
                    float minX = 0f, minY = 0f, maxX = 0f, maxY = 0f;
                    var first = true;
                    foreach (var lx in new[] { 0f, textW })
                    {
                        foreach (var ly in new[] { -ascent, descent })
                        {
                            var rx = lx * cos - ly * sin;
                            var ry = lx * sin + ly * cos;
                            if (first) { minX = maxX = rx; minY = maxY = ry; first = false; }
                            else
                            {
                                if (rx < minX) minX = rx; if (rx > maxX) maxX = rx;
                                if (ry < minY) minY = ry; if (ry > maxY) maxY = ry;
                            }
                        }
                    }

                    float startX, cellCornerX;
                    switch (style.Alignment.Horizontal)
                    {
                        case HorizontalAlignment.Center:
                        case HorizontalAlignment.CenterContinuous:
                            if (isClockwiseDiagonal)
                            {
                                startX = maxX;
                                cellCornerX = rect.Left + padding + rect.Width * 0.30f;
                            }
                            else
                            {
                                startX = (minX + maxX) / 2f;
                                cellCornerX = rect.MidX;
                            }
                            break;
                        case HorizontalAlignment.Right:
                            startX = maxX; cellCornerX = rect.Right - padding;
                            break;
                        default:
                            startX = minX; cellCornerX = rect.Left + padding;
                            break;
                    }

                    // Horizontally the box is pinned by its edge (so leaning glyphs stay inside the cell),
                    // but vertically Excel puts the first character's baseline on the reading edge and lets
                    // ascenders/descenders spill past it - bottom edge for counter-clockwise text, top for
                    // clockwise.
                    var cellCornerY = ccwDeg >= 0 ? rect.Bottom - vpad : rect.Top + vpad;

                    anchorX = cellCornerX - startX;
                    anchorY = cellCornerY;
                    tx = 0f;
                    baseline = 0f;
                }

                canvas.Save();
                // Clip to the (possibly sheared) cell outline so rotated text is trimmed to the same
                // parallelogram the border draws, not to the upright grid rectangle.
                if (shear != 0f)
                {
                    var dxTop = -shear * rect.Height;
                    using (var clip = new SKPath())
                    {
                        clip.MoveTo(rect.Left + dxTop, rect.Top);
                        clip.LineTo(rect.Right + dxTop, rect.Top);
                        clip.LineTo(rect.Right, rect.Bottom);
                        clip.LineTo(rect.Left, rect.Bottom);
                        clip.Close();
                        canvas.ClipPath(clip, SKClipOperation.Intersect, true);
                    }
                }
                canvas.Translate(anchorX, anchorY);
                canvas.RotateDegrees(canvasDeg);
                DrawRuns(canvas, fontContext, text, tx, baseline, paint, style.Font);
                canvas.Restore();
            }
        }

        /// <summary>
        /// Draws vertically-stacked cell text (textRotation 255): each character is upright and placed
        /// below the previous one, the block aligned within the cell.
        /// </summary>
        private void DrawStackedCellText(SKCanvas canvas, SKRect rect, string text, CellFontContext fontContext, StyleValue style)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var color = _context.Colors.Resolve(style.Font.Color, SKColors.Black);
            var padding = (float)SheetLayout.HorizontalPaddingPt;
            var glyphAdvance = StackedGlyphAdvancePt(fontContext);
            var visibleLength = 0;
            for (var i = 0; i < text.Length; i++)
            {
                if (text[i] != ' ')
                {
                    visibleLength++;
                }
            }

            if (visibleLength == 0)
            {
                return;
            }

            var availableHeight = rect.Height - 2f * (float)SheetLayout.VerticalPaddingPt;
            var maxRowsPerColumn = availableHeight > 0f
                ? Math.Max(1, (int)Math.Floor((availableHeight + glyphAdvance * 0.15f) / glyphAdvance))
                : visibleLength;
            if (maxRowsPerColumn < 1)
            {
                maxRowsPerColumn = 1;
            }

            var columnCount = (visibleLength + maxRowsPerColumn - 1) / maxRowsPerColumn;
            var rowsInFirstColumn = Math.Min(visibleLength, maxRowsPerColumn);
            var blockHeight = rowsInFirstColumn * glyphAdvance;
            var columnAdvance = StackedColumnAdvancePt(fontContext);

            float top;
            switch (style.Alignment.Vertical)
            {
                case VerticalAlignment.Top:
                    top = rect.Top + (float)SheetLayout.VerticalPaddingPt;
                    break;
                case VerticalAlignment.Center:
                    top = rect.Top + (rect.Height - blockHeight) / 2f;
                    break;
                default:
                    top = rect.Bottom - (float)SheetLayout.VerticalPaddingPt - blockHeight;
                    break;
            }

            var blockWidth = columnCount * columnAdvance;
            float left;
            switch (style.Alignment.Horizontal)
            {
                case HorizontalAlignment.Left:
                    left = rect.Left + padding;
                    break;
                case HorizontalAlignment.Right:
                    left = rect.Right - padding - blockWidth;
                    break;
                default:
                    left = rect.MidX - blockWidth / 2f;
                    break;
            }

            using (var paint = new SKPaint { Color = color, IsAntialias = true })
            {
                var metricsFont = new SKFont(_context.Fonts.Resolve(style.Font), fontContext.SizePt);
                var metrics = metricsFont.Metrics;
                metricsFont.Dispose();
                var ascent = -metrics.Ascent;

                canvas.Save();
                canvas.ClipRect(rect);
                var visibleIndex = 0;
                for (var i = 0; i < text.Length; i++)
                {
                    var ch = text.Substring(i, 1);
                    if (ch == " ")
                    {
                        continue;
                    }

                    var w = fontContext.Measure(ch);
                    var columnIndex = visibleIndex / maxRowsPerColumn;
                    var rowIndex = visibleIndex % maxRowsPerColumn;
                    var visualColumnIndex = columnCount - 1 - columnIndex;
                    var centerX = left + visualColumnIndex * columnAdvance + columnAdvance / 2f;
                    var baseline = top + rowIndex * glyphAdvance + ascent;
                    DrawRuns(canvas, fontContext, ch, centerX - w / 2f, baseline, paint, style.Font);
                    visibleIndex++;
                }
                canvas.Restore();
            }
        }

        private static float StackedGlyphAdvancePt(CellFontContext fontContext)
        {
            var textHeight = fontContext.TextHeightPt;
            if (textHeight <= 0f)
            {
                return fontContext.LineHeightPt;
            }

            return textHeight + Math.Max(0.6f, fontContext.SizePt * 0.08f);
        }

        private static float StackedColumnAdvancePt(CellFontContext fontContext)
        {
            var zeroWidth = fontContext.Measure("0");
            return Math.Max(fontContext.SizePt * 0.95f, zeroWidth + Math.Max(0.8f, fontContext.SizePt * 0.1f));
        }

        /// <summary>
        /// True when the cell holds a plain number displayed with the General format, which is the
        /// only case where Excel adapts the number of shown digits to the column width.
        /// </summary>
        private static bool IsGeneralNumeric(object value, StyleValue style)
        {
            if (!IsPlainNumber(value))
            {
                return false;
            }

            var formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);
            return string.IsNullOrWhiteSpace(formatCode) || string.Equals(formatCode, "General", StringComparison.Ordinal);
        }

        private static bool IsPlainNumber(object value)
        {
            return value is double || value is float || value is decimal
                || value is int || value is long || value is short || value is byte;
        }

        /// <summary>
        /// Formats a number under Excel's General rules. Excel caps General display at 11 characters
        /// (including sign, decimal point, and exponent) and additionally reduces precision to fit
        /// the cell width, so precision is dropped until both limits are satisfied.
        /// </summary>
        private string FitGeneralNumber(object value, float maxWidthPt, CellFontContext font)
        {
            const int GeneralMaxChars = 11;
            var number = ToDouble(value);

            for (var significantDigits = 11; significantDigits >= 1; significantDigits--)
            {
                var candidate = number.ToString("G" + significantDigits.ToString(System.Globalization.CultureInfo.InvariantCulture), System.Globalization.CultureInfo.InvariantCulture);
                if (candidate.Length <= GeneralMaxChars && font.Measure(candidate) <= maxWidthPt)
                {
                    return candidate;
                }
            }

            // Nothing fits (extremely narrow cell); the caller clips the least-precise form.
            return number.ToString("G1", System.Globalization.CultureInfo.InvariantCulture);
        }

        private static double ToDouble(object value)
        {
            if (value is double) return (double)value;
            if (value is float) return (float)value;
            if (value is decimal) return (double)(decimal)value;
            if (value is int) return (int)value;
            if (value is long) return (long)value;
            if (value is short) return (short)value;
            if (value is byte) return (byte)value;
            return System.Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string StripLineBreaks(string text)
        {
            if (text.IndexOf('\n') < 0 && text.IndexOf('\r') < 0)
            {
                return text;
            }

            return text.Replace("\r\n", string.Empty).Replace("\r", string.Empty).Replace("\n", string.Empty);
        }

        private static HorizontalAlignment ResolveHorizontal(HorizontalAlignment declared, CellRecord record)
        {
            if (declared != HorizontalAlignment.General)
            {
                return declared;
            }

            switch (record.Kind)
            {
                case CellValueKind.Number:
                case CellValueKind.DateTime:
                    return HorizontalAlignment.Right;
                case CellValueKind.Boolean:
                case CellValueKind.Error:
                    return HorizontalAlignment.Center;
                default:
                    if (record.Value is string)
                    {
                        return HorizontalAlignment.Left;
                    }

                    // Formula results / untyped numerics: right-align numbers, left-align the rest.
                    if (DisplayTextFormatterSupport.IsNumericValue(record.Value) || record.Value is DateTime)
                    {
                        return HorizontalAlignment.Right;
                    }

                    return HorizontalAlignment.Left;
            }
        }
    }
}
