using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Resolved geometry for one worksheet: per-column widths and per-row heights in points, their
    /// cumulative offsets, the used cell range, and a merge lookup. Produced once, then consumed by
    /// pagination and rendering.
    /// </summary>
    internal sealed class SheetLayout
    {
        /// <summary>Horizontal text padding inside a cell, per side, in points (~2px at 96 DPI).</summary>
        public const double HorizontalPaddingPt = 1.5d;

        /// <summary>Vertical text padding inside a cell, per side, in points.</summary>
        public const double VerticalPaddingPt = 1.0d;
        public const double AutoFitSingleLineVerticalPaddingPt = 0.72d;

        private const double DefaultColumnWidthChars = 8.43d;
        private const double DefaultRowHeightPt = 15d;

        public WorksheetModel Sheet;

        public int FirstRow;
        public int FirstColumn;
        public int LastRow;
        public int LastColumn;

        public double[] ColumnWidthPt;   // indexed 0..LastColumn
        public double[] RowHeightPt;      // indexed 0..LastRow
        public bool[] ColumnHidden;
        public bool[] RowHidden;

        public double[] ColumnStartPt;    // length LastColumn+2, cumulative left edge (0-based origin)
        public double[] RowStartPt;       // length LastRow+2, cumulative top edge

        private Dictionary<long, MergeRegion> _mergeTopLeft;
        private HashSet<long> _mergeCovered;

        public bool IsEmpty { get { return LastRow < FirstRow || LastColumn < FirstColumn; } }

        public static SheetLayout Build(RenderContext context, WorksheetModel sheet)
        {
            var layout = new SheetLayout();
            layout.Sheet = sheet;
            layout.ComputeUsedRange(sheet);
            layout.BuildMergeLookup(sheet);
            layout.ComputeColumns(context, sheet);
            layout.ComputeRows(context, sheet);
            layout.ComputeCumulativeOffsets();
            layout.TrimFloatingObjectAnchorOvershoot(sheet);
            return layout;
        }

        private void ComputeUsedRange(WorksheetModel sheet)
        {
            int minRow = int.MaxValue, minCol = int.MaxValue, maxRow = -1, maxCol = -1;

            foreach (var pair in sheet.Cells)
            {
                var address = pair.Key;
                if (pair.Value != null && pair.Value.Kind == CellValueKind.Blank && pair.Value.Value == null
                    && IsDefaultStyle(pair.Value.Style))
                {
                    continue;
                }

                if (address.RowIndex < minRow) minRow = address.RowIndex;
                if (address.ColumnIndex < minCol) minCol = address.ColumnIndex;
                if (address.RowIndex > maxRow) maxRow = address.RowIndex;
                if (address.ColumnIndex > maxCol) maxCol = address.ColumnIndex;
            }

            foreach (var merge in sheet.MergeRegions)
            {
                if (merge.FirstRow < minRow) minRow = merge.FirstRow;
                if (merge.FirstColumn < minCol) minCol = merge.FirstColumn;
                var lastR = merge.FirstRow + Math.Max(1, merge.TotalRows) - 1;
                var lastC = merge.FirstColumn + Math.Max(1, merge.TotalColumns) - 1;
                if (lastR > maxRow) maxRow = lastR;
                if (lastC > maxCol) maxCol = lastC;
            }

            // Anchored pictures extend the printable area (Excel paginates to include them), so a
            // picture below/right of the data still gets its own page range.
            foreach (var picture in sheet.Pictures)
            {
                if (picture.UpperLeftRow < minRow) minRow = picture.UpperLeftRow;
                if (picture.UpperLeftColumn < minCol) minCol = picture.UpperLeftColumn;
                if (picture.LowerRightRow > maxRow) maxRow = picture.LowerRightRow;
                if (picture.LowerRightColumn > maxCol) maxCol = picture.LowerRightColumn;
            }

            // Charts extend the printable area the same way as pictures.
            foreach (var chart in sheet.Charts)
            {
                if (chart.UpperLeftRow < minRow) minRow = chart.UpperLeftRow;
                if (chart.UpperLeftColumn < minCol) minCol = chart.UpperLeftColumn;
                if (chart.LowerRightRow > maxRow) maxRow = chart.LowerRightRow;
                if (chart.LowerRightColumn > maxCol) maxCol = chart.LowerRightColumn;
            }

            // Shapes (autoShapes, connectors) extend the printable area too.
            foreach (var shape in sheet.Shapes)
            {
                if (shape.UpperLeftRow < minRow) minRow = shape.UpperLeftRow;
                if (shape.UpperLeftColumn < minCol) minCol = shape.UpperLeftColumn;
                if (shape.LowerRightRow > maxRow) maxRow = shape.LowerRightRow;
                if (shape.LowerRightColumn > maxCol) maxCol = shape.LowerRightColumn;
            }

            // SmartArt diagrams extend the printable area too.
            foreach (var smartArt in sheet.SmartArts)
            {
                if (smartArt.UpperLeftRow < minRow) minRow = smartArt.UpperLeftRow;
                if (smartArt.UpperLeftColumn < minCol) minCol = smartArt.UpperLeftColumn;
                if (smartArt.LowerRightRow > maxRow) maxRow = smartArt.LowerRightRow;
                if (smartArt.LowerRightColumn > maxCol) maxCol = smartArt.LowerRightColumn;
            }

            if (maxRow < 0 || maxCol < 0)
            {
                // Nothing to render: keep an empty range anchored at A1.
                FirstRow = 0; FirstColumn = 0; LastRow = -1; LastColumn = -1;
                return;
            }

            // Keep the printable coordinate system anchored at A1, matching Excel's pagination:
            // leading empty rows/columns still consume printable space even when they contain no
            // visible content.
            FirstRow = 0;
            FirstColumn = 0;
            LastRow = maxRow;
            LastColumn = maxCol;
        }

        private static bool IsDefaultStyle(StyleValue style)
        {
            return style == null
                || (style.Pattern == FillPatternKind.None
                    && style.Borders.Left.Style == BorderStyle.None
                    && style.Borders.Right.Style == BorderStyle.None
                    && style.Borders.Top.Style == BorderStyle.None
                    && style.Borders.Bottom.Style == BorderStyle.None);
        }

        private void BuildMergeLookup(WorksheetModel sheet)
        {
            _mergeTopLeft = new Dictionary<long, MergeRegion>();
            _mergeCovered = new HashSet<long>();

            foreach (var merge in sheet.MergeRegions)
            {
                var rows = Math.Max(1, merge.TotalRows);
                var cols = Math.Max(1, merge.TotalColumns);
                _mergeTopLeft[Key(merge.FirstRow, merge.FirstColumn)] = merge;

                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        if (r == 0 && c == 0)
                        {
                            continue;
                        }

                        _mergeCovered.Add(Key(merge.FirstRow + r, merge.FirstColumn + c));
                    }
                }
            }
        }

        private void ComputeColumns(RenderContext context, WorksheetModel sheet)
        {
            if (IsEmpty)
            {
                ColumnWidthPt = new double[0];
                ColumnHidden = new bool[0];
                return;
            }

            var count = LastColumn + 1;
            ColumnWidthPt = new double[count];
            ColumnHidden = new bool[count];

            // An explicit defaultColWidth is a stored (padded) width; the fabricated fallback (8.43) is
            // a display-character count, so each uses the matching conversion to match Excel.
            var defaultPt = sheet.DefaultColumnWidth.HasValue && sheet.DefaultColumnWidth.Value > 0
                ? RenderUnits.ColumnWidthCharsToPoints(sheet.DefaultColumnWidth.Value, context.MaxDigitWidthPixels)
                : RenderUnits.DefaultColumnWidthCharsToPoints(DefaultColumnWidthChars, context.MaxDigitWidthPixels);

            for (var c = 0; c < count; c++)
            {
                ColumnWidthPt[c] = defaultPt;
            }

            foreach (var range in sheet.Columns)
            {
                var min = Math.Max(0, range.MinColumnIndex);
                var max = Math.Min(LastColumn, range.MaxColumnIndex);
                for (var c = min; c <= max; c++)
                {
                    if (range.Hidden)
                    {
                        ColumnHidden[c] = true;
                        ColumnWidthPt[c] = 0d;
                    }
                    else if (range.Width.HasValue && range.Width.Value > 0)
                    {
                        ColumnWidthPt[c] = RenderUnits.ColumnWidthCharsToPoints(range.Width.Value, context.MaxDigitWidthPixels);
                    }
                }
            }
        }

        private void ComputeRows(RenderContext context, WorksheetModel sheet)
        {
            if (IsEmpty)
            {
                RowHeightPt = new double[0];
                RowHidden = new bool[0];
                return;
            }

            var count = LastRow + 1;
            RowHeightPt = new double[count];
            RowHidden = new bool[count];
            var rowHeightScale = ResolveStoredRowHeightScale(context, sheet);

            // Excel only treats the stored defaultRowHeight as authoritative when customHeight is
            // set; otherwise it re-derives the row height from the workbook's default font (a stored
            // value can be stale after the default font changed). Match that so CJK-font sheets get
            // the correct, shorter rows instead of a leftover Latin-font height.
            var customHeightSet = sheet.CustomHeight.HasValue && sheet.CustomHeight.Value;
            double defaultHeight;
            if (customHeightSet && sheet.DefaultRowHeight.HasValue && sheet.DefaultRowHeight.Value > 0)
            {
                defaultHeight = sheet.DefaultRowHeight.Value;
            }
            else if (context.FontDerivedRowHeightPt > 0d)
            {
                defaultHeight = context.FontDerivedRowHeightPt;
            }
            else
            {
                defaultHeight = sheet.DefaultRowHeight.HasValue && sheet.DefaultRowHeight.Value > 0
                    ? sheet.DefaultRowHeight.Value
                    : DefaultRowHeightPt;
            }

            for (var r = 0; r < count; r++)
            {
                RowModel rowModel;
                var hasRow = sheet.Rows.TryGetValue(r, out rowModel) && rowModel != null;

                if (hasRow && rowModel.Hidden)
                {
                    RowHidden[r] = true;
                    RowHeightPt[r] = 0d;
                    continue;
                }

                if (hasRow && rowModel.Height.HasValue && rowModel.Height.Value > 0 && rowModel.CustomHeight)
                {
                    // A true customHeight row is user-authored and stays authoritative.
                    RowHeightPt[r] = rowModel.Height.Value;
                }
                else if (hasRow && rowModel.Height.HasValue && rowModel.Height.Value > 0 && RowUsesCachedAutoHeight(sheet, r))
                {
                    // Wrap/rotation rows often persist Excel's latest auto-fit result only in the
                    // row height cache. Reuse that cached value for rotated text instead of
                    // re-fitting with our simpler text model, which can otherwise greatly
                    // overestimate vertical text.
                    RowHeightPt[r] = rowModel.Height.Value;
                }
                else
                {
                    // A stored height without customHeight is only Excel's cached auto-fit result.
                    // Re-fit from the actual row content so stale heights shrink back to the same
                    // size Excel uses during PDF export.
                    var fittedHeight = Math.Max(defaultHeight, AutoRowHeight(context, sheet, r));
                    if (hasRow && rowModel.Height.HasValue && rowModel.Height.Value > 0)
                    {
                        fittedHeight = Math.Max(fittedHeight, Math.Min(rowModel.Height.Value * rowHeightScale, fittedHeight));
                    }

                    RowHeightPt[r] = fittedHeight;
                }
            }
        }

        /// <summary>
        /// Some workbooks keep a stale defaultRowHeight after the default font changed. Excel's PDF
        /// export visually scales explicit row heights by the ratio between the current font-derived
        /// default row height and that stale stored height; apply the same normalization so custom
        /// rows stay visually consistent with the active font.
        /// </summary>
        private static double ResolveStoredRowHeightScale(RenderContext context, WorksheetModel sheet)
        {
            if (sheet == null || !sheet.DefaultRowHeight.HasValue || sheet.DefaultRowHeight.Value <= 0d)
            {
                return 1d;
            }

            if (sheet.CustomHeight.HasValue && sheet.CustomHeight.Value)
            {
                return 1d;
            }

            if (context == null || context.FontDerivedRowHeightPt <= 0d)
            {
                return 1d;
            }

            var scale = context.FontDerivedRowHeightPt / sheet.DefaultRowHeight.Value;
            if (scale < 0.75d || scale > 1.25d)
            {
                return 1d;
            }

            return scale;
        }

        /// <summary>
        /// True when a non-custom-height row contains rotated text and Excel's stored row height
        /// acts as the latest cached auto-fit result worth preserving.
        /// </summary>
        private bool RowUsesCachedAutoHeight(WorksheetModel sheet, int row)
        {
            for (var col = 0; col <= LastColumn; col++)
            {
                if (ColumnHidden[col])
                {
                    continue;
                }

                CellRecord record;
                if (!sheet.Cells.TryGetValue(new CellAddress(row, col), out record) || record == null)
                {
                    continue;
                }

                if (IsMergeCovered(row, col) || SpansMultipleRows(row, col))
                {
                    continue;
                }

                var style = record.Style ?? StyleValue.Default;
                if (style.Alignment.TextRotation > 0)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Computes the height a row needs to fit wrap-text cells whose height is not fixed.
        /// </summary>
        private double AutoRowHeight(RenderContext context, WorksheetModel sheet, int row)
        {
            double needed = 0d;

            for (var col = 0; col <= LastColumn; col++)
            {
                if (ColumnHidden[col])
                {
                    continue;
                }

                CellRecord record;
                if (!sheet.Cells.TryGetValue(new CellAddress(row, col), out record) || record == null)
                {
                    continue;
                }

                var style = record.Style ?? StyleValue.Default;
                var rotation = style.Alignment.TextRotation;
                var isRotated = rotation > 0;
                // Skip cells that are merged across rows; their height is governed by the merge.
                if (IsMergeCovered(row, col) || SpansMultipleRows(row, col))
                {
                    continue;
                }

                var text = DisplayTextFormatter.FormatDisplayValue(record.Value, style, context.Culture);
                if (string.IsNullOrEmpty(text))
                {
                    continue;
                }

                var font = context.GetFontContext(style.Font);

                if (isRotated)
                {
                    // Excel fits the row tightly to the rotated text's bounding box (no cell padding),
                    // so the angled text can graze the top/bottom borders as it does in Excel.
                    var rotatedHeight = RotatedTextHeightPt(font, text, rotation);
                    if (rotatedHeight > needed)
                    {
                        needed = rotatedHeight;
                    }
                    continue;
                }

                var lineCount = 1;
                if (style.Alignment.WrapText)
                {
                    var availableWidth = MergedWidthPt(row, col) - 2d * HorizontalPaddingPt;
                    availableWidth = EffectiveWrapWidthPt(availableWidth, style.Alignment);
                    if (availableWidth <= 0d)
                    {
                        continue;
                    }

                    var lines = font.WrapLines(text, (float)availableWidth);
                    lineCount = Math.Max(1, lines.Count);
                }

                if (!style.Alignment.WrapText && style.Alignment.Horizontal == HorizontalAlignment.Distributed)
                {
                    // Excel's distributed horizontal alignment keeps single-line text visually on
                    // one baseline but auto-fits the row to about two default line boxes.
                    lineCount = Math.Max(lineCount, 2);
                }

                var verticalPadding = style.Alignment.WrapText
                    ? 2d * VerticalPaddingPt
                    : 2d * AutoFitSingleLineVerticalPaddingPt;
                var blockHeight = lineCount * font.LineHeightPt + verticalPadding;
                if (blockHeight > needed)
                {
                    needed = blockHeight;
                }
            }

            return needed;
        }

        /// <summary>
        /// Height (excluding padding) that a single line of rotated cell text occupies. Angled text
        /// (1-180) contributes the height of its rotated bounding box; vertically-stacked text (255)
        /// contributes one line per character.
        /// </summary>
        private static double RotatedTextHeightPt(CellFontContext font, string text, int rotation)
        {
            var flat = text.Replace("\r", string.Empty).Replace("\n", string.Empty);
            if (rotation >= 255)
            {
                return Math.Max(1, flat.Length) * font.LineHeightPt;
            }

            var ccwDeg = rotation <= 90 ? rotation : rotation - 90;
            var rad = ccwDeg * Math.PI / 180d;

            // Size the row to the rotated text's tight ink box (what Excel fits to), not the font's
            // nominal line box, so a row of rotated text is not left several points too tall.
            float inkWidth, inkHeight;
            font.MeasureInkBounds(flat, out inkWidth, out inkHeight);
            return inkWidth * Math.Abs(Math.Sin(rad)) + inkHeight * Math.Abs(Math.Cos(rad));
        }

        private static double EffectiveWrapWidthPt(double availableWidth, AlignmentValue alignment)
        {
            if (availableWidth <= 0d || alignment == null)
            {
                return availableWidth;
            }

            if (alignment.Horizontal == HorizontalAlignment.Distributed)
            {
                var indentUnits = alignment.IndentLevel + Math.Max(0, alignment.RelativeIndent);
                if (indentUnits > 0)
                {
                    return availableWidth / (1d + 2d * indentUnits);
                }
            }

            return availableWidth;
        }

        /// <summary>
        /// True when any non-merged-continuation cell in the row carries a text-rotation style, so a
        /// cached (non-customHeight) row height must be re-fitted rather than trusted.
        /// </summary>
        private bool RowHasRotatedText(WorksheetModel sheet, int row)
        {
            for (var col = 0; col <= LastColumn; col++)
            {
                if (ColumnHidden[col])
                {
                    continue;
                }

                CellRecord record;
                if (!sheet.Cells.TryGetValue(new CellAddress(row, col), out record) || record == null)
                {
                    continue;
                }

                var style = record.Style ?? StyleValue.Default;
                if (style.Alignment.TextRotation > 0 && !IsMergeCovered(row, col) && !SpansMultipleRows(row, col))
                {
                    return true;
                }
            }

            return false;
        }

        private bool SpansMultipleRows(int row, int col)
        {
            MergeRegion merge;
            if (_mergeTopLeft.TryGetValue(Key(row, col), out merge))
            {
                return Math.Max(1, merge.TotalRows) > 1;
            }

            return false;
        }

        private double MergedWidthPt(int row, int col)
        {
            MergeRegion merge;
            if (_mergeTopLeft.TryGetValue(Key(row, col), out merge))
            {
                double width = 0d;
                var cols = Math.Max(1, merge.TotalColumns);
                for (var c = 0; c < cols && col + c <= LastColumn; c++)
                {
                    width += ColumnWidthPt[col + c];
                }

                return width;
            }

            return col <= LastColumn ? ColumnWidthPt[col] : 0d;
        }

        private void ComputeCumulativeOffsets()
        {
            if (IsEmpty)
            {
                ColumnStartPt = new double[] { 0d };
                RowStartPt = new double[] { 0d };
                return;
            }

            ColumnStartPt = new double[LastColumn + 2];
            for (var c = 0; c <= LastColumn; c++)
            {
                ColumnStartPt[c + 1] = ColumnStartPt[c] + ColumnWidthPt[c];
            }

            RowStartPt = new double[LastRow + 2];
            for (var r = 0; r <= LastRow; r++)
            {
                RowStartPt[r + 1] = RowStartPt[r] + RowHeightPt[r];
            }
        }

        /// <summary>
        /// Some real files keep coarse two-cell anchors whose lower-right cell extends beyond the
        /// actual object size, while the drawing's EMU extent carries the real width/height. Trim the
        /// printable used range back to the actual floating-object bounds so pagination does not keep
        /// extra blank trailing columns/rows solely because of that anchor overshoot.
        /// </summary>
        private void TrimFloatingObjectAnchorOvershoot(WorksheetModel sheet)
        {
            if (IsEmpty)
            {
                return;
            }

            var maxRow = -1;
            var maxCol = -1;

            foreach (var pair in sheet.Cells)
            {
                var address = pair.Key;
                if (pair.Value != null && pair.Value.Kind == CellValueKind.Blank && pair.Value.Value == null
                    && IsDefaultStyle(pair.Value.Style))
                {
                    continue;
                }

                if (address.RowIndex > maxRow)
                {
                    maxRow = address.RowIndex;
                }

                if (address.ColumnIndex > maxCol)
                {
                    maxCol = address.ColumnIndex;
                }
            }

            foreach (var merge in sheet.MergeRegions)
            {
                var lastRow = merge.FirstRow + Math.Max(1, merge.TotalRows) - 1;
                var lastCol = merge.FirstColumn + Math.Max(1, merge.TotalColumns) - 1;
                if (lastRow > maxRow)
                {
                    maxRow = lastRow;
                }

                if (lastCol > maxCol)
                {
                    maxCol = lastCol;
                }
            }

            foreach (var picture in sheet.Pictures)
            {
                if (picture == null)
                {
                    continue;
                }

                if (picture.ExtentCx > 0)
                {
                    var right = ColumnEdgePt(picture.UpperLeftColumn, picture.UpperLeftColumnOffset) + picture.ExtentCx / 12700d;
                    var effectiveLastColumn = FindLastCoveredColumn(right);
                    if (effectiveLastColumn > maxCol)
                    {
                        maxCol = effectiveLastColumn;
                    }
                }

                if (picture.ExtentCy > 0)
                {
                    var bottom = RowEdgePt(picture.UpperLeftRow, picture.UpperLeftRowOffset) + picture.ExtentCy / 12700d;
                    var effectiveLastRow = FindLastCoveredRow(bottom);
                    if (effectiveLastRow > maxRow)
                    {
                        maxRow = effectiveLastRow;
                    }
                }
            }

            foreach (var chart in sheet.Charts)
            {
                if (chart == null)
                {
                    continue;
                }

                if (chart.LowerRightRow > maxRow)
                {
                    maxRow = chart.LowerRightRow;
                }

                if (chart.LowerRightColumn > maxCol)
                {
                    maxCol = chart.LowerRightColumn;
                }
            }

            foreach (var shape in sheet.Shapes)
            {
                if (shape == null)
                {
                    continue;
                }

                if (shape.LowerRightRow > maxRow)
                {
                    maxRow = shape.LowerRightRow;
                }

                if (shape.LowerRightColumn > maxCol)
                {
                    maxCol = shape.LowerRightColumn;
                }
            }

            foreach (var smartArt in sheet.SmartArts)
            {
                if (smartArt == null)
                {
                    continue;
                }

                if (smartArt.LowerRightRow > maxRow)
                {
                    maxRow = smartArt.LowerRightRow;
                }

                if (smartArt.LowerRightColumn > maxCol)
                {
                    maxCol = smartArt.LowerRightColumn;
                }
            }

            if (maxRow < 0 || maxCol < 0)
            {
                return;
            }

            LastRow = maxRow;
            LastColumn = maxCol;
        }

        private double ColumnEdgePt(int column, long emuOffset)
        {
            if (column < 0)
            {
                column = 0;
            }

            if (column > ColumnStartPt.Length - 1)
            {
                column = ColumnStartPt.Length - 1;
            }

            return ColumnStartPt[column] + emuOffset / 12700d;
        }

        private double RowEdgePt(int row, long emuOffset)
        {
            if (row < 0)
            {
                row = 0;
            }

            if (row > RowStartPt.Length - 1)
            {
                row = RowStartPt.Length - 1;
            }

            return RowStartPt[row] + emuOffset / 12700d;
        }

        private int FindLastCoveredColumn(double rightPt)
        {
            for (var column = 0; column <= LastColumn; column++)
            {
                if (ColumnStartPt[column + 1] >= rightPt - 0.01d)
                {
                    return column;
                }
            }

            return LastColumn;
        }

        private int FindLastCoveredRow(double bottomPt)
        {
            for (var row = 0; row <= LastRow; row++)
            {
                if (RowStartPt[row + 1] >= bottomPt - 0.01d)
                {
                    return row;
                }
            }

            return LastRow;
        }

        public bool TryGetPictureColumnSpan(PictureModel picture, out int startColumn, out int endColumn, out double widthPt)
        {
            startColumn = 0;
            endColumn = 0;
            widthPt = 0d;

            if (picture == null || IsEmpty)
            {
                return false;
            }

            startColumn = Math.Max(FirstColumn, picture.UpperLeftColumn);
            if (startColumn > LastColumn)
            {
                startColumn = LastColumn;
            }

            var left = ColumnEdgePt(picture.UpperLeftColumn, picture.UpperLeftColumnOffset);
            double right;
            if (picture.ExtentCx > 0)
            {
                right = left + picture.ExtentCx / 12700d;
            }
            else
            {
                right = ColumnEdgePt(picture.LowerRightColumn, picture.LowerRightColumnOffset);
            }

            if (right <= left + 0.01d)
            {
                return false;
            }

            endColumn = FindLastCoveredColumn(right);
            if (endColumn < startColumn)
            {
                endColumn = startColumn;
            }

            widthPt = right - left;
            return true;
        }

        public bool IsMergeCovered(int row, int col)
        {
            return _mergeCovered.Contains(Key(row, col));
        }

        public bool TryGetMergeOrigin(int row, int col, out MergeRegion merge)
        {
            return _mergeTopLeft.TryGetValue(Key(row, col), out merge);
        }

        private static long Key(int row, int col)
        {
            return ((long)row << 20) | (uint)col;
        }
    }
}
