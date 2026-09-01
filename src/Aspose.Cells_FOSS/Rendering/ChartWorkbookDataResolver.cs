using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartWorkbookDataResolver
    {
        public static List<double?> ResolveNumericRange(WorkbookModel workbook, string formula, DateSystem dateSystem)
        {
            var result = new List<double?>();
            WorksheetModel sheet;
            int firstRow;
            int firstColumn;
            int lastRow;
            int lastColumn;
            if (!TryParseSingleRangeReference(workbook, formula, out sheet, out firstRow, out firstColumn, out lastRow, out lastColumn))
            {
                return result;
            }

            if (firstColumn == lastColumn)
            {
                for (var row = firstRow; row <= lastRow; row++)
                {
                    result.Add(ReadNumericCell(sheet, row, firstColumn, dateSystem));
                }
                return result;
            }

            if (firstRow == lastRow)
            {
                for (var column = firstColumn; column <= lastColumn; column++)
                {
                    result.Add(ReadNumericCell(sheet, firstRow, column, dateSystem));
                }
            }

            return result;
        }

        public static List<string> ResolveCategoryLabels(WorkbookModel workbook, string formula, string formatCode, DateSystem dateSystem, CultureInfo culture)
        {
            var numericValues = ResolveNumericRange(workbook, formula, dateSystem);
            var result = new List<string>(numericValues.Count);
            for (var i = 0; i < numericValues.Count; i++)
            {
                if (numericValues[i].HasValue)
                {
                    result.Add(ChartXmlParser.FormatLabel(numericValues[i].Value, formatCode, dateSystem, culture));
                }
                else
                {
                    result.Add(string.Empty);
                }
            }

            return result;
        }

        public static List<string> ResolveStringRange(WorkbookModel workbook, string formula)
        {
            var result = new List<string>();
            WorksheetModel sheet;
            int firstRow;
            int firstColumn;
            int lastRow;
            int lastColumn;
            if (!TryParseSingleRangeReference(workbook, formula, out sheet, out firstRow, out firstColumn, out lastRow, out lastColumn))
            {
                return result;
            }

            if (firstColumn == lastColumn)
            {
                for (var row = firstRow; row <= lastRow; row++)
                {
                    result.Add(ReadStringCell(workbook, sheet, row, firstColumn));
                }
                return result;
            }

            if (firstRow == lastRow)
            {
                for (var column = firstColumn; column <= lastColumn; column++)
                {
                    result.Add(ReadStringCell(workbook, sheet, firstRow, column));
                }
            }

            return result;
        }

        public static string ResolveSingleString(WorkbookModel workbook, string formula)
        {
            var values = ResolveStringRange(workbook, formula);
            if (values.Count > 0)
            {
                return values[0];
            }

            return string.Empty;
        }

        private static double? ReadNumericCell(WorksheetModel sheet, int row, int column, DateSystem dateSystem)
        {
            CellRecord record;
            if (!sheet.Cells.TryGetValue(new CellAddress(row, column), out record) || record == null || record.Value == null)
            {
                return null;
            }

            if (record.Value is DateTime)
            {
                return DateSerialConverter.ToSerial((DateTime)record.Value, dateSystem);
            }

            if (record.Value is double)
            {
                return (double)record.Value;
            }

            if (record.Value is int)
            {
                return (int)record.Value;
            }

            if (record.Value is long)
            {
                return (long)record.Value;
            }

            if (record.Value is float)
            {
                return (float)record.Value;
            }

            if (record.Value is decimal)
            {
                return (double)(decimal)record.Value;
            }

            double parsed;
            if (double.TryParse(record.Value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
            {
                return parsed;
            }

            return null;
        }

        private static string ReadStringCell(WorkbookModel workbook, WorksheetModel sheet, int row, int column)
        {
            CellRecord record;
            if (!sheet.Cells.TryGetValue(new CellAddress(row, column), out record) || record == null || record.Value == null)
            {
                return string.Empty;
            }

            var style = record.Style != null ? record.Style : StyleValue.Default;
            var culture = workbook != null && workbook.Settings != null
                ? workbook.Settings.DisplayCulture
                : CultureInfo.InvariantCulture;
            return DisplayTextFormatter.FormatDisplayValue(record.Value, style, culture);
        }

        private static bool TryParseSingleRangeReference(WorkbookModel workbook, string formula, out WorksheetModel sheet, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
        {
            sheet = null;
            firstRow = 0;
            firstColumn = 0;
            lastRow = 0;
            lastColumn = 0;

            if (workbook == null || string.IsNullOrEmpty(formula))
            {
                return false;
            }

            var bang = formula.IndexOf('!');
            if (bang <= 0 || bang >= formula.Length - 1)
            {
                return false;
            }

            var sheetName = formula.Substring(0, bang);
            var range = formula.Substring(bang + 1);
            sheetName = UnescapeSheetName(sheetName);

            sheet = FindWorksheet(workbook, sheetName);
            if (sheet == null)
            {
                return false;
            }

            var parts = range.Split(':');
            if (parts.Length == 1)
            {
                parts = new[] { parts[0], parts[0] };
            }

            CellAddress first;
            CellAddress last;
            if (!TryParseAbsoluteCell(parts[0], out first) || !TryParseAbsoluteCell(parts[1], out last))
            {
                return false;
            }

            firstRow = Math.Min(first.RowIndex, last.RowIndex);
            firstColumn = Math.Min(first.ColumnIndex, last.ColumnIndex);
            lastRow = Math.Max(first.RowIndex, last.RowIndex);
            lastColumn = Math.Max(first.ColumnIndex, last.ColumnIndex);
            return true;
        }

        private static WorksheetModel FindWorksheet(WorkbookModel workbook, string sheetName)
        {
            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (string.Equals(workbook.Worksheets[i].Name, sheetName, StringComparison.Ordinal))
                {
                    return workbook.Worksheets[i];
                }
            }

            return null;
        }

        private static string UnescapeSheetName(string sheetName)
        {
            var trimmed = sheetName.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '\'' && trimmed[trimmed.Length - 1] == '\'')
            {
                trimmed = trimmed.Substring(1, trimmed.Length - 2).Replace("''", "'");
            }

            return trimmed;
        }

        private static bool TryParseAbsoluteCell(string value, out CellAddress address)
        {
            var normalized = value.Replace("$", string.Empty).Trim();
            try
            {
                address = CellAddress.Parse(normalized);
                return true;
            }
            catch (Exception)
            {
                address = default(CellAddress);
                return false;
            }
        }
    }
}
