using System;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class TableStyleResolver
    {
        private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly ColorValue HeaderFill = new ColorValue(255, 0, 78, 110);
        private static readonly ColorValue DataFill = new ColorValue(255, 255, 255, 255);
        private static readonly ColorValue StripeFill = new ColorValue(255, 192, 230, 245);
        private static readonly ColorValue BorderFill = new ColorValue(255, 159, 192, 205);
        private static readonly ColorValue LightText = new ColorValue(255, 255, 255, 255);
        private readonly WorkbookModel _workbook;
        private readonly XElement _customStyles;

        public TableStyleResolver()
            : this(null)
        {
        }

        public TableStyleResolver(WorkbookModel workbook)
        {
            _workbook = workbook;
            if (workbook != null && !string.IsNullOrEmpty(workbook.RawTableStylesXml))
            {
                try
                {
                    _customStyles = XElement.Parse(workbook.RawTableStylesXml);
                }
                catch (Exception)
                {
                    _customStyles = null;
                }
            }
        }

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

        private StyleValue ApplyTableStyle(ListObjectModel table, int row, StyleValue baseStyle)
        {
            var type = ListObjectSupport.TableStyleTypeFromName(table.TableStyleName);
            if (type == TableStyleType.Custom)
            {
                return ApplyCustomTableStyle(table, row, baseStyle);
            }

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

        private StyleValue ApplyCustomTableStyle(ListObjectModel table, int row, StyleValue baseStyle)
        {
            if (_customStyles == null || _workbook == null)
            {
                return baseStyle;
            }

            XElement definition = null;
            foreach (var candidate in _customStyles.Elements(Main + "tableStyle"))
            {
                if (string.Equals((string)candidate.Attribute("name"), table.TableStyleName, StringComparison.Ordinal))
                {
                    definition = candidate;
                    break;
                }
            }

            if (definition == null)
            {
                return baseStyle;
            }

            var result = baseStyle.Clone();
            ApplyElement(definition, "wholeTable", result);
            if (table.ShowHeaderRow && row == table.StartRow)
            {
                ApplyElement(definition, "headerRow", result);
            }
            else if (table.ShowTotals && row == table.EndRow)
            {
                ApplyElement(definition, "totalRow", result);
            }
            else if (table.ShowRowStripes)
            {
                var firstDataRow = table.StartRow + (table.ShowHeaderRow ? 1 : 0);
                if ((row - firstDataRow) % 2 == 0)
                {
                    ApplyElement(definition, "firstRowStripe", result);
                }
                else
                {
                    ApplyElement(definition, "secondRowStripe", result);
                }
            }

            return result;
        }

        private void ApplyElement(XElement definition, string type, StyleValue result)
        {
            foreach (var element in definition.Elements(Main + "tableStyleElement"))
            {
                if (!string.Equals((string)element.Attribute("type"), type, StringComparison.Ordinal))
                {
                    continue;
                }

                int index;
                if (!int.TryParse((string)element.Attribute("dxfId"), NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out index)
                    || index < 0 || index >= _workbook.TableDifferentialStyles.Count
                    || index >= _workbook.RawTableDifferentialStylesXml.Count)
                {
                    return;
                }

                XElement raw;
                try
                {
                    raw = XElement.Parse(_workbook.RawTableDifferentialStylesXml[index]);
                }
                catch (Exception)
                {
                    return;
                }

                var value = _workbook.TableDifferentialStyles[index];
                var font = raw.Element(Main + "font");
                if (font != null)
                {
                    var bold = font.Element(Main + "b");
                    if (bold != null)
                    {
                        result.Font.Bold = !string.Equals((string)bold.Attribute("val"), "0", StringComparison.Ordinal);
                    }

                    if (font.Element(Main + "color") != null)
                    {
                        result.Font.Color = value.Font.Color;
                    }
                }

                var fill = raw.Element(Main + "fill");
                var pattern = fill != null ? fill.Element(Main + "patternFill") : null;
                if (pattern != null)
                {
                    result.Pattern = value.Pattern;
                    result.ForegroundColor = value.ForegroundColor;
                    result.BackgroundColor = value.BackgroundColor;
                }

                if (raw.Element(Main + "border") != null)
                {
                    result.Borders = value.Borders.Clone();
                }

                return;
            }
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
