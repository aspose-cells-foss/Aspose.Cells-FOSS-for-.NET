using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ConditionalFillRunRenderer
    {
        public static int Draw(SKCanvas canvas, PageLayout page, ConditionalFormatEvaluator formatting,
            TableStyleResolver tableStyles)
        {
            var layout = page.Sheet;
            var count = 0;
            using (var paint = new SKPaint { Style = SKPaintStyle.Fill, IsAntialias = false })
            {
                for (var column = page.StartColumn; column <= page.EndColumn; column++)
                {
                    var firstRow = -1;
                    var runColor = SKColors.Transparent;
                    for (var row = page.StartRow; row <= page.EndRow + 1; row++)
                    {
                        var color = SKColors.Transparent;
                        var hasFill = row <= page.EndRow
                            && TryGetFill(layout, formatting, tableStyles, row, column, out color);
                        if (firstRow >= 0 && (!hasFill || !color.Equals(runColor)))
                        {
                            var left = (float)(layout.ColumnStartPt[column]
                                - layout.ColumnStartPt[page.StartColumn]);
                            var right = (float)(layout.ColumnStartPt[column + 1]
                                - layout.ColumnStartPt[page.StartColumn]);
                            var top = (float)(layout.RowStartPt[firstRow]
                                - layout.RowStartPt[page.StartRow]);
                            var bottom = (float)(layout.RowStartPt[row]
                                - layout.RowStartPt[page.StartRow]);
                            paint.Color = runColor;
                            canvas.DrawRect(new SKRect(left, top, right, bottom), paint);
                            count++;
                            firstRow = -1;
                        }

                        if (hasFill && firstRow < 0)
                        {
                            firstRow = row;
                            runColor = color;
                        }
                    }
                }
            }

            return count;
        }

        private static bool TryGetFill(SheetLayout layout, ConditionalFormatEvaluator formatting,
            TableStyleResolver tableStyles, int row, int column, out SKColor color)
        {
            color = SKColors.Transparent;
            if (layout.RowHidden[row] || layout.ColumnHidden[column])
            {
                return false;
            }

            MergeRegion merge;
            if (layout.IsMergeCovered(row, column)
                || layout.TryGetMergeOrigin(row, column, out merge))
            {
                return false;
            }

            ConditionalFormatEvaluator.CellEffect effect;
            if (!formatting.TryGet(row, column, out effect) || !effect.HasFill
                || effect.Fill.Alpha == 0)
            {
                return false;
            }

            CellRecord record;
            layout.Sheet.Cells.TryGetValue(new CellAddress(row, column), out record);
            var baseStyle = record != null && record.Style != null ? record.Style : StyleValue.Default;
            var style = tableStyles.Resolve(layout.Sheet, row, column, baseStyle);
            if (style != null && style.Alignment != null
                && style.Alignment.TextRotation > 90 && style.Alignment.TextRotation < 180)
            {
                return false;
            }

            color = effect.Fill;
            return true;
        }
    }
}
