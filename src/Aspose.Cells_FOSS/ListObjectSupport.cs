using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Internal helpers for ListObject validation and model construction.
    /// </summary>
    internal static class ListObjectSupport
    {
        internal static void ValidateRange(int startRow, int startColumn, int endRow, int endColumn)
        {
            if (startRow < 0)
            {
                throw new CellsException("Table startRow must be non-negative.");
            }

            if (startColumn < 0)
            {
                throw new CellsException("Table startColumn must be non-negative.");
            }

            if (endRow < startRow)
            {
                throw new CellsException("Table endRow must be greater than or equal to startRow.");
            }

            if (endColumn < startColumn)
            {
                throw new CellsException("Table endColumn must be greater than or equal to startColumn.");
            }
        }

        internal static void ValidateDisplayName(string displayName)
        {
            if (string.IsNullOrEmpty(displayName))
            {
                throw new CellsException("Table DisplayName must be non-empty.");
            }

            var first = displayName[0];
            if (!char.IsLetter(first) && first != '_' && first != '\\')
            {
                throw new CellsException("Table DisplayName '" + displayName + "' must start with a letter or underscore.");
            }
        }

        internal static void ValidateUniqueDisplayName(IReadOnlyList<ListObjectModel> existing, string displayName, int skipIndex)
        {
            for (var i = 0; i < existing.Count; i++)
            {
                if (i == skipIndex)
                {
                    continue;
                }

                if (string.Equals(existing[i].DisplayName, displayName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new CellsException("A table with display name '" + displayName + "' already exists in this worksheet.");
                }
            }
        }

        internal static void ValidateNoOverlap(IReadOnlyList<ListObjectModel> existing, int startRow, int startColumn, int endRow, int endColumn, int skipIndex)
        {
            for (var i = 0; i < existing.Count; i++)
            {
                if (i == skipIndex)
                {
                    continue;
                }

                var other = existing[i];
                if (RangesOverlap(startRow, startColumn, endRow, endColumn, other.StartRow, other.StartColumn, other.EndRow, other.EndColumn))
                {
                    throw new CellsException("The table range overlaps with existing table '" + other.DisplayName + "'.");
                }
            }
        }

        internal static bool RangesOverlap(int r1s, int c1s, int r1e, int c1e, int r2s, int c2s, int r2e, int c2e)
        {
            return r1s <= r2e && r2s <= r1e && c1s <= c2e && c2s <= c1e;
        }

        internal static ListObjectModel CreateModel(WorksheetModel worksheetModel, int startRow, int startColumn, int endRow, int endColumn, bool hasHeaders, int tableNumber)
        {
            var displayName = "Table" + tableNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var model = new ListObjectModel
            {
                DisplayName = displayName,
                Name = displayName,
                StartRow = startRow,
                StartColumn = startColumn,
                EndRow = endRow,
                EndColumn = endColumn,
                ShowHeaderRow = hasHeaders,
                HasAutoFilter = hasHeaders,
            };

            var columnCount = endColumn - startColumn + 1;
            for (var c = 0; c < columnCount; c++)
            {
                var columnName = ResolveColumnName(worksheetModel, startRow, startColumn + c, hasHeaders, c + 1);
                model.Columns.Add(new ListColumnModel(c + 1, columnName));
            }

            return model;
        }

        internal static void RebuildColumns(ListObjectModel model, WorksheetModel worksheetModel)
        {
            model.Columns.Clear();
            var columnCount = model.EndColumn - model.StartColumn + 1;
            for (var c = 0; c < columnCount; c++)
            {
                var columnName = ResolveColumnName(worksheetModel, model.StartRow, model.StartColumn + c, model.ShowHeaderRow, c + 1);
                model.Columns.Add(new ListColumnModel(c + 1, columnName));
            }
        }

        private static string ResolveColumnName(WorksheetModel worksheetModel, int headerRow, int columnIndex, bool hasHeaders, int defaultColumnNumber)
        {
            if (hasHeaders)
            {
                var address = new CellAddress(headerRow, columnIndex);
                CellRecord record;
                if (worksheetModel.Cells.TryGetValue(address, out record) && record.Value is string)
                {
                    var name = (string)record.Value;
                    if (!string.IsNullOrEmpty(name))
                    {
                        return name;
                    }
                }
            }

            return "Column" + defaultColumnNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        internal static string TableStyleTypeToName(TableStyleType type)
        {
            if (type == TableStyleType.None || type == TableStyleType.Custom)
            {
                return string.Empty;
            }

            return type.ToString();
        }

        internal static TableStyleType TableStyleTypeFromName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return TableStyleType.None;
            }

            TableStyleType parsed;
            if (TryParseTableStyleType(name, out parsed))
            {
                return parsed;
            }

            return TableStyleType.Custom;
        }

        private static bool TryParseTableStyleType(string name, out TableStyleType result)
        {
            try
            {
                result = (TableStyleType)Enum.Parse(typeof(TableStyleType), name, ignoreCase: false);
                return result != TableStyleType.None && result != TableStyleType.Custom;
            }
            catch (ArgumentException)
            {
                result = TableStyleType.None;
                return false;
            }
        }

        internal static string SanitizeDisplayName(string displayName)
        {
            if (string.IsNullOrEmpty(displayName))
            {
                return displayName;
            }

            return displayName.Replace(" ", "_");
        }
    }
}
