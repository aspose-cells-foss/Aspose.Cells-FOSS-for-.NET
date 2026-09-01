using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class TableStyleResolver
    {
        private static readonly ColorValue HeaderFill = new ColorValue(255, 0, 78, 110);
        private static readonly ColorValue DataFill = new ColorValue(255, 255, 255, 255);
        private static readonly ColorValue StripeFill = new ColorValue(255, 192, 230, 245);
        private static readonly ColorValue BorderFill = new ColorValue(255, 159, 192, 205);
        private static readonly ColorValue LightText = new ColorValue(255, 255, 255, 255);

        public StyleValue Resolve(WorksheetModel sheet, int row, int column, StyleValue baseStyle)
        {
            if (sheet == null || sheet.ListObjects == null || sheet.ListObjects.Count == 0)
            {
                return baseStyle ?? StyleValue.Default;
            }

            var style = baseStyle ?? StyleValue.Default;
            for (var i = 0; i < sheet.ListObjects.Count; i++)
            {
                var table = sheet.ListObjects[i];
                if (!Contains(table, row, column))
                {
                    continue;
                }

                return ApplyTableStyle(table, row, style);
            }

            return style;
        }

        private static bool Contains(ListObjectModel table, int row, int column)
        {
            if (table == null)
            {
                return false;
            }

            return row >= table.StartRow
                && row <= table.EndRow
                && column >= table.StartColumn
                && column <= table.EndColumn;
        }

        private static StyleValue ApplyTableStyle(ListObjectModel table, int row, StyleValue baseStyle)
        {
            var type = ListObjectSupport.TableStyleTypeFromName(table.TableStyleName);
            if (type != TableStyleType.TableStyleMedium2)
            {
                return baseStyle;
            }

            if (table.ShowHeaderRow && row == table.StartRow)
            {
                var styled = baseStyle.Clone();
                styled.Pattern = FillPatternKind.Solid;
                styled.ForegroundColor = HeaderFill;
                styled.BackgroundColor = HeaderFill;
                styled.Font.Bold = true;
                styled.Font.Color = LightText;
                ApplyBorders(styled, LightText);
                return styled;
            }

            if (table.ShowTotals && row == table.EndRow)
            {
                var totals = baseStyle.Clone();
                totals.Pattern = FillPatternKind.Solid;
                totals.ForegroundColor = StripeFill;
                totals.BackgroundColor = StripeFill;
                totals.Font.Bold = true;
                ApplyBorders(totals, BorderFill);
                return totals;
            }

            var data = baseStyle.Clone();
            data.Pattern = FillPatternKind.Solid;
            data.ForegroundColor = DataFill;
            data.BackgroundColor = DataFill;
            ApplyBorders(data, BorderFill);

            if (table.ShowRowStripes)
            {
                var dataStartRow = table.StartRow + (table.ShowHeaderRow ? 1 : 0);
                var dataEndRow = table.EndRow - (table.ShowTotals ? 1 : 0);
                if (row >= dataStartRow && row <= dataEndRow)
                {
                    var stripeIndex = row - dataStartRow;
                    if ((stripeIndex % 2) == 0)
                    {
                        var striped = data.Clone();
                        striped.ForegroundColor = StripeFill;
                        striped.BackgroundColor = StripeFill;
                        return striped;
                    }
                }
            }

            return data;
        }

        private static void ApplyBorders(StyleValue style, ColorValue color)
        {
            style.Borders.Left.Style = BorderStyle.Thin;
            style.Borders.Right.Style = BorderStyle.Thin;
            style.Borders.Top.Style = BorderStyle.Thin;
            style.Borders.Bottom.Style = BorderStyle.Thin;

            style.Borders.Left.Color = color;
            style.Borders.Right.Color = color;
            style.Borders.Top.Color = color;
            style.Borders.Bottom.Color = color;
        }
    }
}
