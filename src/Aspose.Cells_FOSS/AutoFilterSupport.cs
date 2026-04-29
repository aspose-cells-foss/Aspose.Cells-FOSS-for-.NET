using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    internal static class AutoFilterSupport
    {
        internal static string NormalizeOptionalRange(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string normalized;
            if (!TryNormalizeRange(value, out normalized))
            {
                throw new CellsException(parameterName + " must be a valid cell or range reference.");
            }

            return normalized;
        }

        internal static string NormalizeRequiredRange(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new CellsException(parameterName + " must be a valid cell or range reference.");
            }

            string normalized;
            if (!TryNormalizeRange(value, out normalized))
            {
                throw new CellsException(parameterName + " must be a valid cell or range reference.");
            }

            return normalized;
        }

        internal static bool TryNormalizeRange(string value, out string normalized)
        {
            normalized = string.Empty;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            MergeRegion region;
            if (!XlsxWorkbookSerializerCommon.TryParseMergeReference(value.Trim(), out region))
            {
                return false;
            }

            normalized = XlsxWorkbookSerializerCommon.ToRangeReference(region);
            return true;
        }

        internal static string NormalizeText(string value, string parameterName)
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            return value;
        }

        internal static string NormalizeOptionalText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value.Trim();
        }

        internal static FilterOperatorType ParseOperatorOrDefault(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return FilterOperatorType.Equal;
            }

            switch (value.Trim())
            {
                case "lessThan":
                    return FilterOperatorType.LessThan;
                case "lessThanOrEqual":
                    return FilterOperatorType.LessOrEqual;
                case "notEqual":
                    return FilterOperatorType.NotEqual;
                case "greaterThanOrEqual":
                    return FilterOperatorType.GreaterOrEqual;
                case "greaterThan":
                    return FilterOperatorType.GreaterThan;
                default:
                    return FilterOperatorType.Equal;
            }
        }

        internal static bool TryParseOperator(string value, out FilterOperatorType operatorType)
        {
            operatorType = FilterOperatorType.Equal;
            if (string.IsNullOrWhiteSpace(value))
            {
                return true;
            }

            switch (value.Trim())
            {
                case "lessThan":
                    operatorType = FilterOperatorType.LessThan;
                    return true;
                case "lessThanOrEqual":
                    operatorType = FilterOperatorType.LessOrEqual;
                    return true;
                case "notEqual":
                    operatorType = FilterOperatorType.NotEqual;
                    return true;
                case "greaterThanOrEqual":
                    operatorType = FilterOperatorType.GreaterOrEqual;
                    return true;
                case "greaterThan":
                    operatorType = FilterOperatorType.GreaterThan;
                    return true;
                case "equal":
                    operatorType = FilterOperatorType.Equal;
                    return true;
                default:
                    return false;
            }
        }

        internal static string ToOperatorName(FilterOperatorType operatorType)
        {
            switch (operatorType)
            {
                case FilterOperatorType.LessThan:
                    return "lessThan";
                case FilterOperatorType.LessOrEqual:
                    return "lessThanOrEqual";
                case FilterOperatorType.NotEqual:
                    return "notEqual";
                case FilterOperatorType.GreaterOrEqual:
                    return "greaterThanOrEqual";
                case FilterOperatorType.GreaterThan:
                    return "greaterThan";
                default:
                    return null;
            }
        }

        internal static int CompareFilterColumns(FilterColumnModel left, FilterColumnModel right)
        {
            return left.ColumnIndex.CompareTo(right.ColumnIndex);
        }
    }
}
