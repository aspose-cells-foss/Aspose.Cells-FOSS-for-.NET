using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;
using static Aspose.Cells_FOSS.XlsxWorkbookConditionalFormatting;
using static Aspose.Cells_FOSS.XlsxWorkbookHyperlinks;
using static Aspose.Cells_FOSS.XlsxWorkbookStyles;
using static Aspose.Cells_FOSS.XlsxWorkbookPageSetup;
using static Aspose.Cells_FOSS.XlsxWorkbookValidations;
using static Aspose.Cells_FOSS.XlsxWorkbookWorksheetProtection;
using static Aspose.Cells_FOSS.XlsxWorkbookAutoFilter;
using static Aspose.Cells_FOSS.XlsxWorkbookWorksheetViews;
using static Aspose.Cells_FOSS.XlsxWorkbookTables;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookWorksheetWriter
    {
        internal static XDocument BuildWorksheet(WorksheetModel worksheet, StyleValue workbookDefaultStyle, Aspose.Cells_FOSS.Core.DateSystem dateSystem, SharedStringRepository sharedStrings, SaveOptions options, StylesheetSaveContext stylesheet, int externalHyperlinkCount, bool hasPictures, bool hasComments)
        {
            var persistedCells = CollectPersistedCells(worksheet, workbookDefaultStyle);
            var worksheetElement = new XElement(MainNs + "worksheet",
                new XAttribute(XNamespace.Xmlns + "r", RelationshipNs));

            var sheetProperties = BuildWorksheetSheetProperties(worksheet);
            if (sheetProperties != null)
            {
                worksheetElement.Add(sheetProperties);
            }

            var dimensionReference = CalculateDimensionReference(persistedCells, worksheet.MergeRegions);
            if (!string.IsNullOrEmpty(dimensionReference))
            {
                worksheetElement.Add(new XElement(MainNs + "dimension", new XAttribute("ref", dimensionReference)));
            }

            var sheetViews = BuildWorksheetViewsElement(worksheet);
            if (sheetViews != null)
            {
                worksheetElement.Add(sheetViews);
            }
            var normalizedColumns = NormalizeColumnRanges(worksheet.Columns);
            if (normalizedColumns.Count > 0)
            {
                worksheetElement.Add(new XElement(MainNs + "cols", BuildColumnElements(normalizedColumns)));
            }

            var sheetData = new XElement(MainNs + "sheetData");
            foreach (var rowIndex in GetWorksheetRowIndexes(persistedCells, worksheet.Rows))
            {
                var rowCells = GetRowCells(persistedCells, rowIndex);
                RowModel rowModel;
                worksheet.Rows.TryGetValue(rowIndex, out rowModel);
                if (rowCells.Count == 0 && !HasRowMetadata(rowModel))
                {
                    continue;
                }

                sheetData.Add(BuildRowElement(rowIndex, rowModel, rowCells, dateSystem, sharedStrings, options, stylesheet));
            }

            worksheetElement.Add(sheetData);

            var sheetProtection = BuildSheetProtectionElement(worksheet);
            if (sheetProtection != null)
            {
                worksheetElement.Add(sheetProtection);
            }

            var autoFilter = BuildAutoFilterElement(worksheet, stylesheet.DifferentialFormatCount);
            if (autoFilter != null)
            {
                worksheetElement.Add(autoFilter);
            }

            var mergeRegions = SortMergeRegions(worksheet.MergeRegions);
            if (mergeRegions.Count > 0)
            {
                worksheetElement.Add(new XElement(MainNs + "mergeCells",
                    new XAttribute("count", mergeRegions.Count),
                    BuildMergeCellElements(mergeRegions)));
            }

            worksheetElement.Add(BuildConditionalFormattingElements(worksheet, stylesheet));

            var dataValidations = BuildDataValidationsElement(worksheet);
            if (dataValidations != null)
            {
                worksheetElement.Add(dataValidations);
            }

            var hyperlinks = BuildHyperlinksElement(worksheet);
            if (hyperlinks != null)
            {
                worksheetElement.Add(hyperlinks);
            }

            var printOptions = BuildPrintOptionsElement(worksheet.PageSetup);
            if (printOptions != null)
            {
                worksheetElement.Add(printOptions);
            }

            var pageMargins = BuildPageMarginsElement(worksheet.PageSetup);
            if (pageMargins != null)
            {
                worksheetElement.Add(pageMargins);
            }

            var pageSetup = BuildPageSetupElement(worksheet.PageSetup);
            if (pageSetup != null)
            {
                worksheetElement.Add(pageSetup);
            }

            var headerFooter = BuildHeaderFooterElement(worksheet.PageSetup);
            if (headerFooter != null)
            {
                worksheetElement.Add(headerFooter);
            }

            var rowBreaks = BuildRowBreaksElement(worksheet.PageSetup);
            if (rowBreaks != null)
            {
                worksheetElement.Add(rowBreaks);
            }

            var columnBreaks = BuildColumnBreaksElement(worksheet.PageSetup);
            if (columnBreaks != null)
            {
                worksheetElement.Add(columnBreaks);
            }

            if (hasPictures)
            {
                var drawingRId = externalHyperlinkCount + worksheet.ListObjects.Count + 1;
                worksheetElement.Add(new XElement(MainNs + "drawing",
                    new XAttribute(RelationshipNs + "id", "rId" + drawingRId.ToString(CultureInfo.InvariantCulture))));
            }

            if (hasComments)
            {
                var vmlRId = externalHyperlinkCount + worksheet.ListObjects.Count + (hasPictures ? 1 : 0) + 2;
                worksheetElement.Add(new XElement(MainNs + "legacyDrawing",
                    new XAttribute(RelationshipNs + "id", "rId" + vmlRId.ToString(CultureInfo.InvariantCulture))));
            }

            var tableRIdStart = externalHyperlinkCount + 1;
            var tableParts = BuildTablePartsElement(worksheet, tableRIdStart);
            if (tableParts != null)
            {
                worksheetElement.Add(tableParts);
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), worksheetElement);
        }

        private static List<KeyValuePair<CellAddress, CellRecord>> CollectPersistedCells(WorksheetModel worksheet, StyleValue workbookDefaultStyle)
        {
            var persistedCells = new List<KeyValuePair<CellAddress, CellRecord>>();
            foreach (var pair in worksheet.Cells)
            {
                if (ShouldPersistCell(workbookDefaultStyle, pair.Value))
                {
                    persistedCells.Add(pair);
                }
            }

            persistedCells.Sort(ComparePersistedCells);
            return persistedCells;
        }

        private static int ComparePersistedCells(KeyValuePair<CellAddress, CellRecord> left, KeyValuePair<CellAddress, CellRecord> right)
        {
            var rowComparison = left.Key.RowIndex.CompareTo(right.Key.RowIndex);
            if (rowComparison != 0)
            {
                return rowComparison;
            }

            return left.Key.ColumnIndex.CompareTo(right.Key.ColumnIndex);
        }

        private static List<XElement> BuildColumnElements(IReadOnlyList<ColumnRangeModel> columns)
        {
            var columnElements = new List<XElement>(columns.Count);
            for (var index = 0; index < columns.Count; index++)
            {
                columnElements.Add(BuildColumnElement(columns[index]));
            }

            return columnElements;
        }

        private static List<KeyValuePair<CellAddress, CellRecord>> GetRowCells(IReadOnlyList<KeyValuePair<CellAddress, CellRecord>> persistedCells, int rowIndex)
        {
            var rowCells = new List<KeyValuePair<CellAddress, CellRecord>>();
            for (var index = 0; index < persistedCells.Count; index++)
            {
                var pair = persistedCells[index];
                if (pair.Key.RowIndex == rowIndex)
                {
                    rowCells.Add(pair);
                }
            }

            return rowCells;
        }

        private static List<MergeRegion> SortMergeRegions(IReadOnlyList<MergeRegion> mergeRegions)
        {
            var sortedRegions = new List<MergeRegion>(mergeRegions.Count);
            for (var index = 0; index < mergeRegions.Count; index++)
            {
                sortedRegions.Add(mergeRegions[index]);
            }

            sortedRegions.Sort(CompareMergeRegions);
            return sortedRegions;
        }

        private static int CompareMergeRegions(MergeRegion left, MergeRegion right)
        {
            var rowComparison = left.FirstRow.CompareTo(right.FirstRow);
            if (rowComparison != 0)
            {
                return rowComparison;
            }

            var columnComparison = left.FirstColumn.CompareTo(right.FirstColumn);
            if (columnComparison != 0)
            {
                return columnComparison;
            }

            var rowCountComparison = left.TotalRows.CompareTo(right.TotalRows);
            if (rowCountComparison != 0)
            {
                return rowCountComparison;
            }

            return left.TotalColumns.CompareTo(right.TotalColumns);
        }

        private static List<XElement> BuildMergeCellElements(IReadOnlyList<MergeRegion> mergeRegions)
        {
            var mergeCellElements = new List<XElement>(mergeRegions.Count);
            for (var index = 0; index < mergeRegions.Count; index++)
            {
                mergeCellElements.Add(new XElement(MainNs + "mergeCell", new XAttribute("ref", ToRangeReference(mergeRegions[index]))));
            }

            return mergeCellElements;
        }

        internal static XElement BuildRowElement(int rowIndex, RowModel rowModel, IReadOnlyList<KeyValuePair<CellAddress, CellRecord>> rowCells, Aspose.Cells_FOSS.Core.DateSystem dateSystem, SharedStringRepository sharedStrings, SaveOptions options, StylesheetSaveContext stylesheet)
        {
            var row = new XElement(MainNs + "row", new XAttribute("r", rowIndex + 1));
            if (rowModel != null && rowModel.Height.HasValue)
            {
                row.SetAttributeValue("ht", rowModel.Height.Value.ToString("R", CultureInfo.InvariantCulture));
                row.SetAttributeValue("customHeight", 1);
            }

            if (rowModel != null && rowModel.Hidden)
            {
                row.SetAttributeValue("hidden", 1);
            }

            if (rowModel != null && rowModel.StyleIndex.HasValue && rowModel.StyleIndex.Value >= 0)
            {
                row.SetAttributeValue("s", rowModel.StyleIndex.Value);
                row.SetAttributeValue("customFormat", 1);
            }

            foreach (var pair in rowCells)
            {
                row.Add(BuildCell(pair.Key, pair.Value, dateSystem, sharedStrings, options, stylesheet));
            }

            return row;
        }

        internal static XElement BuildColumnElement(ColumnRangeModel column)
        {
            var element = new XElement(MainNs + "col",
                new XAttribute("min", column.MinColumnIndex + 1),
                new XAttribute("max", column.MaxColumnIndex + 1));

            if (column.Width.HasValue)
            {
                element.SetAttributeValue("width", column.Width.Value.ToString("R", CultureInfo.InvariantCulture));
                element.SetAttributeValue("customWidth", 1);
            }

            if (column.Hidden)
            {
                element.SetAttributeValue("hidden", 1);
            }

            if (column.StyleIndex.HasValue && column.StyleIndex.Value >= 0)
            {
                element.SetAttributeValue("style", column.StyleIndex.Value);
            }

            return element;
        }

        internal static bool HasRowMetadata(RowModel rowModel)
        {
            return rowModel != null && (rowModel.Height.HasValue || rowModel.Hidden || rowModel.StyleIndex.HasValue);
        }

        internal static bool HasColumnMetadata(ColumnRangeModel column)
        {
            return column.Width.HasValue || column.Hidden || column.StyleIndex.HasValue;
        }

        internal static List<int> GetWorksheetRowIndexes(IReadOnlyList<KeyValuePair<CellAddress, CellRecord>> persistedCells, IReadOnlyDictionary<int, RowModel> rows)
        {
            var indexes = new HashSet<int>();
            foreach (var pair in persistedCells)
            {
                indexes.Add(pair.Key.RowIndex);
            }

            foreach (var rowIndex in rows.Keys)
            {
                indexes.Add(rowIndex);
            }

            var orderedIndexes = new List<int>(indexes);
            orderedIndexes.Sort();
            return orderedIndexes;
        }

        internal static List<ColumnRangeModel> NormalizeColumnRanges(IReadOnlyList<ColumnRangeModel> columns)
        {
            var ordered = new List<ColumnRangeModel>();
            for (var index = 0; index < columns.Count; index++)
            {
                var column = columns[index];
                if (!HasColumnMetadata(column))
                {
                    continue;
                }

                ordered.Add(new ColumnRangeModel
                {
                    MinColumnIndex = column.MinColumnIndex,
                    MaxColumnIndex = column.MaxColumnIndex,
                    Width = column.Width,
                    Hidden = column.Hidden,
                    StyleIndex = column.StyleIndex,
                });
            }

            ordered.Sort(CompareColumnRangesByBounds);
            if (ordered.Count == 0)
            {
                return ordered;
            }

            var normalized = new List<ColumnRangeModel> { ordered[0] };
            for (var index = 1; index < ordered.Count; index++)
            {
                var current = ordered[index];
                var previous = normalized[normalized.Count - 1];
                if (previous.MaxColumnIndex + 1 >= current.MinColumnIndex && ColumnMetadataEqual(previous, current))
                {
                    previous.MaxColumnIndex = Math.Max(previous.MaxColumnIndex, current.MaxColumnIndex);
                    continue;
                }

                normalized.Add(current);
            }

            return normalized;
        }

        private static int CompareColumnRangesByBounds(ColumnRangeModel left, ColumnRangeModel right)
        {
            var minComparison = left.MinColumnIndex.CompareTo(right.MinColumnIndex);
            if (minComparison != 0)
            {
                return minComparison;
            }

            return left.MaxColumnIndex.CompareTo(right.MaxColumnIndex);
        }

        internal static bool ColumnMetadataEqual(ColumnRangeModel left, ColumnRangeModel right)
        {
            return Nullable.Equals(left.Width, right.Width)
                && left.Hidden == right.Hidden
                && left.StyleIndex == right.StyleIndex;
        }

        internal static string CalculateDimensionReference(IReadOnlyList<KeyValuePair<CellAddress, CellRecord>> persistedCells, IReadOnlyList<MergeRegion> mergeRegions)
        {
            var hasCells = persistedCells.Count > 0;
            var hasMerges = mergeRegions.Count > 0;
            if (!hasCells && !hasMerges)
            {
                return string.Empty;
            }

            var minRow = int.MaxValue;
            var minColumn = int.MaxValue;
            var maxRow = int.MinValue;
            var maxColumn = int.MinValue;

            foreach (var pair in persistedCells)
            {
                minRow = Math.Min(minRow, pair.Key.RowIndex);
                minColumn = Math.Min(minColumn, pair.Key.ColumnIndex);
                maxRow = Math.Max(maxRow, pair.Key.RowIndex);
                maxColumn = Math.Max(maxColumn, pair.Key.ColumnIndex);
            }

            foreach (var region in mergeRegions)
            {
                minRow = Math.Min(minRow, region.FirstRow);
                minColumn = Math.Min(minColumn, region.FirstColumn);
                maxRow = Math.Max(maxRow, region.FirstRow + region.TotalRows - 1);
                maxColumn = Math.Max(maxColumn, region.FirstColumn + region.TotalColumns - 1);
            }

            var firstAddress = new CellAddress(minRow, minColumn).ToString();
            var lastAddress = new CellAddress(maxRow, maxColumn).ToString();
            return string.Equals(firstAddress, lastAddress, StringComparison.Ordinal) ? firstAddress : firstAddress + ":" + lastAddress;
        }

        internal static XElement BuildCell(CellAddress address, CellRecord record, Aspose.Cells_FOSS.Core.DateSystem dateSystem, SharedStringRepository sharedStrings, SaveOptions options, StylesheetSaveContext stylesheet)
        {
            var cell = new XElement(MainNs + "c", new XAttribute("r", address.ToString()));
            var styleIndex = stylesheet.GetStyleIndex(record);
            if (styleIndex > 0)
            {
                cell.SetAttributeValue("s", styleIndex);
            }
            var hasFormula = !string.IsNullOrEmpty(record.Formula);
            if (hasFormula)
            {
                cell.Add(new XElement(MainNs + "f", record.Formula));
            }

            if (record.Value == null)
            {
                return cell;
            }

            if (record.Kind == CellValueKind.Error)
            {
                cell.SetAttributeValue("t", "e");
                cell.Add(new XElement(MainNs + "v", record.Value.ToString() ?? string.Empty));
                return cell;
            }

            if (record.Value is string)
            {
                var text = (string)record.Value;
                if (hasFormula)
                {
                    cell.SetAttributeValue("t", "str");
                    cell.Add(new XElement(MainNs + "v", text));
                }
                else if (options.UseSharedStrings)
                {
                    cell.SetAttributeValue("t", "s");
                    cell.Add(new XElement(MainNs + "v", sharedStrings.Intern(text).ToString(CultureInfo.InvariantCulture)));
                }
                else
                {
                    cell.SetAttributeValue("t", "inlineStr");
                    cell.Add(new XElement(MainNs + "is", CreateTextElement(text)));
                }
            }
            else if (record.Value is bool)
            {
                var booleanValue = (bool)record.Value;
                cell.SetAttributeValue("t", "b");
                cell.Add(new XElement(MainNs + "v", booleanValue ? "1" : "0"));
            }
            else if (record.Value is DateTime)
            {
                var dateTime = (DateTime)record.Value;
                cell.Add(new XElement(MainNs + "v", DateSerialConverter.ToSerial(dateTime, dateSystem).ToString("R", CultureInfo.InvariantCulture)));
            }
            else if (record.Value is byte)
            {
                cell.Add(new XElement(MainNs + "v", ((byte)record.Value).ToString(CultureInfo.InvariantCulture)));
            }
            else if (record.Value is short)
            {
                cell.Add(new XElement(MainNs + "v", ((short)record.Value).ToString(CultureInfo.InvariantCulture)));
            }
            else if (record.Value is int)
            {
                cell.Add(new XElement(MainNs + "v", ((int)record.Value).ToString(CultureInfo.InvariantCulture)));
            }
            else if (record.Value is long)
            {
                cell.Add(new XElement(MainNs + "v", ((long)record.Value).ToString(CultureInfo.InvariantCulture)));
            }
            else if (record.Value is float)
            {
                cell.Add(new XElement(MainNs + "v", ((float)record.Value).ToString("R", CultureInfo.InvariantCulture)));
            }
            else if (record.Value is double)
            {
                cell.Add(new XElement(MainNs + "v", ((double)record.Value).ToString("R", CultureInfo.InvariantCulture)));
            }
            else if (record.Value is decimal)
            {
                cell.Add(new XElement(MainNs + "v", ((decimal)record.Value).ToString(CultureInfo.InvariantCulture)));
            }
            else
            {
                var formattable = record.Value as IFormattable;
                if (formattable != null)
                {
                    cell.Add(new XElement(MainNs + "v", formattable.ToString(null, CultureInfo.InvariantCulture)));
                }
                else if (hasFormula)
                {
                    cell.SetAttributeValue("t", "str");
                    cell.Add(new XElement(MainNs + "v", record.Value.ToString() ?? string.Empty));
                }
                else if (options.UseSharedStrings)
                {
                    var fallbackText = record.Value.ToString() ?? string.Empty;
                    cell.SetAttributeValue("t", "s");
                    cell.Add(new XElement(MainNs + "v", sharedStrings.Intern(fallbackText).ToString(CultureInfo.InvariantCulture)));
                }
                else
                {
                    cell.SetAttributeValue("t", "inlineStr");
                    cell.Add(new XElement(MainNs + "is", CreateTextElement(record.Value.ToString() ?? string.Empty)));
                }
            }

            return cell;
        }
    }
}
