using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class ConditionalFormatExpressionMatcher
    {
        private static readonly Regex Comparison = new Regex(
            @"\A\s*(\$?[A-Z]{1,3}\$?[1-9][0-9]*)\s*(>=|<=|<>|>|<|=)\s*([+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?)\s*\z",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private readonly WorksheetModel _sheet;
        private readonly int _anchorRow;
        private readonly int _anchorColumn;
        private readonly int _referenceRow;
        private readonly int _referenceColumn;
        private readonly bool _absoluteRow;
        private readonly bool _absoluteColumn;
        private readonly string _operator;
        private readonly double _threshold;

        private ConditionalFormatExpressionMatcher(WorksheetModel sheet, CellArea area,
            CellAddress reference, bool absoluteRow, bool absoluteColumn,
            string comparisonOperator, double threshold)
        {
            _sheet = sheet;
            _anchorRow = area.StartRow;
            _anchorColumn = area.StartColumn;
            _referenceRow = reference.RowIndex;
            _referenceColumn = reference.ColumnIndex;
            _absoluteRow = absoluteRow;
            _absoluteColumn = absoluteColumn;
            _operator = comparisonOperator;
            _threshold = threshold;
        }

        internal static ConditionalFormatExpressionMatcher TryCreate(WorksheetModel sheet,
            CellArea area, string formula)
        {
            var expression = (formula ?? string.Empty).Trim();
            if (expression.StartsWith("=", StringComparison.Ordinal))
            {
                expression = expression.Substring(1);
            }

            var match = Comparison.Match(expression);
            double threshold;
            if (!match.Success || !double.TryParse(match.Groups[3].Value, NumberStyles.Float,
                CultureInfo.InvariantCulture, out threshold))
            {
                return null;
            }

            var token = match.Groups[1].Value;
            var absoluteColumn = token[0] == '$';
            var rowMarker = token.LastIndexOf('$');
            var absoluteRow = rowMarker > 0;
            CellAddress reference;
            try
            {
                reference = CellAddress.Parse(token.Replace("$", string.Empty));
            }
            catch (ArgumentException)
            {
                return null;
            }

            return new ConditionalFormatExpressionMatcher(sheet, area, reference,
                absoluteRow, absoluteColumn, match.Groups[2].Value, threshold);
        }

        internal bool Matches(int row, int column)
        {
            var referencedRow = _absoluteRow ? _referenceRow
                : (long)_referenceRow + row - _anchorRow;
            var referencedColumn = _absoluteColumn ? _referenceColumn
                : (long)_referenceColumn + column - _anchorColumn;
            if (referencedRow < 0 || referencedRow > int.MaxValue
                || referencedColumn < 0 || referencedColumn > int.MaxValue)
            {
                return false;
            }

            CellRecord record;
            _sheet.Cells.TryGetValue(new CellAddress((int)referencedRow,
                (int)referencedColumn), out record);
            double value;
            if (!ConditionalFormatEvaluator.TryNumber(record, out value))
            {
                return false;
            }

            switch (_operator)
            {
                case ">": return value > _threshold;
                case ">=": return value >= _threshold;
                case "<": return value < _threshold;
                case "<=": return value <= _threshold;
                case "=": return value == _threshold;
                case "<>": return value != _threshold;
                default: return false;
            }
        }
    }
}
