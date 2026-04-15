using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookConditionalFormatting;
using static Aspose.Cells_FOSS.XlsxWorkbookHyperlinks;
using static Aspose.Cells_FOSS.XlsxWorkbookStyles;
using static Aspose.Cells_FOSS.XlsxWorkbookDefinedNames;
using static Aspose.Cells_FOSS.XlsxWorkbookPageSetup;
using static Aspose.Cells_FOSS.XlsxWorkbookValidations;
using static Aspose.Cells_FOSS.XlsxWorkbookWorksheetProtection;
using static Aspose.Cells_FOSS.XlsxWorkbookAutoFilter;
using static Aspose.Cells_FOSS.XlsxWorkbookWorksheetViews;
using static Aspose.Cells_FOSS.XlsxWorkbookProperties;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookSerializerCommon
    {
        internal const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        internal const string SharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        internal const string StylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        internal const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        internal const string ExtendedPropertiesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        internal static readonly XNamespace MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal static readonly XNamespace RelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        internal static readonly XNamespace PackageRelationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        internal static readonly XNamespace ContentTypeNs = "http://schemas.openxmlformats.org/package/2006/content-types";
        internal static readonly XNamespace XmlNs = XNamespace.Xml;
        internal static readonly HashSet<int> BuiltInDateFormats = new HashSet<int> { 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47 };

        internal static bool ShouldPersistCell(StyleValue workbookDefaultStyle, CellRecord record)
        {
            return record.IsExplicitlyStored
                || !string.IsNullOrEmpty(record.Formula)
                || record.Value != null
                || record.Kind != CellValueKind.Blank
                || !StylesEqual(record.Style, workbookDefaultStyle);
        }

        internal static void WriteXmlEntry(ZipArchive archive, string path, XDocument document)
        {
            var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            {
                using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    Indent = false,
                    CloseOutput = false,
                }))
                {
                    document.Save(writer);
                }
            }
        }

        internal static XDocument BuildContentTypes(WorkbookModel model, bool hasSharedStrings, bool hasDateStyles, bool hasCoreProperties, bool hasExtendedProperties)
            {
                var root = new XElement(ContentTypeNs + "Types",
                    new XElement(ContentTypeNs + "Default",
                        new XAttribute("Extension", "rels"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(ContentTypeNs + "Default",
                        new XAttribute("Extension", "xml"),
                        new XAttribute("ContentType", "application/xml")),
                    new XElement(ContentTypeNs + "Override",
                        new XAttribute("PartName", "/xl/workbook.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")));

                if (hasCoreProperties)
                {
                    root.Add(new XElement(ContentTypeNs + "Override",
                        new XAttribute("PartName", "/docProps/core.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml")));
                }

                if (hasExtendedProperties)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/docProps/app.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml")));
            }

            for (var index = 0; index < model.Worksheets.Count; index++)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", $"/xl/worksheets/sheet{index + 1}.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            }

            if (hasSharedStrings)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/xl/sharedStrings.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")));
            }

            if (hasDateStyles)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/xl/styles.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")));
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
        }

        internal static XDocument BuildRootRelationships(bool hasCoreProperties, bool hasExtendedProperties)
        {
            var relationships = new XElement(PackageRelationshipNs + "Relationships",
                new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId1"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                    new XAttribute("Target", "xl/workbook.xml")));

            var relationshipId = 2;
            if (hasCoreProperties)
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + relationshipId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", CorePropertiesRelationshipType),
                    new XAttribute("Target", "docProps/core.xml")));
                relationshipId++;
            }

            if (hasExtendedProperties)
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + relationshipId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", ExtendedPropertiesRelationshipType),
                    new XAttribute("Target", "docProps/app.xml")));
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), relationships);
        }

        internal static XDocument BuildWorkbook(WorkbookModel model)
        {
            var workbook = new XElement(MainNs + "workbook",
                new XAttribute(XNamespace.Xmlns + "r", RelationshipNs));

            var workbookProperties = BuildWorkbookPropertiesElement(model);
            if (workbookProperties != null)
            {
                workbook.Add(workbookProperties);
            }

            var workbookProtection = BuildWorkbookProtectionElement(model);
            if (workbookProtection != null)
            {
                workbook.Add(workbookProtection);
            }

            var bookViews = BuildBookViewsElement(model);
            if (bookViews != null)
            {
                workbook.Add(bookViews);
            }

            var sheets = new XElement(MainNs + "sheets");
            for (var index = 0; index < model.Worksheets.Count; index++)
            {
                var worksheet = model.Worksheets[index];
                var sheet = new XElement(MainNs + "sheet",
                    new XAttribute("name", worksheet.Name),
                    new XAttribute("sheetId", index + 1),
                    new XAttribute(RelationshipNs + "id", $"rId{index + 1}"));

                if (worksheet.Visibility == SheetVisibility.Hidden)
                {
                    sheet.SetAttributeValue("state", "hidden");
                }
                else if (worksheet.Visibility == SheetVisibility.VeryHidden)
                {
                    sheet.SetAttributeValue("state", "veryHidden");
                }

                sheets.Add(sheet);
            }

            workbook.Add(sheets);

            var definedNames = BuildDefinedNames(model);
            if (definedNames != null)
            {
                workbook.Add(definedNames);
            }

            var calculationProperties = BuildCalculationPropertiesElement(model);
            if (calculationProperties != null)
            {
                workbook.Add(calculationProperties);
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), workbook);
        }
        internal static XDocument BuildWorkbookRelationships(WorkbookModel model, bool hasSharedStrings, bool hasDateStyles)
        {
            var relationships = new XElement(PackageRelationshipNs + "Relationships");
            var relationshipId = 1;

            for (var index = 0; index < model.Worksheets.Count; index++)
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", $"rId{relationshipId++}"),
                    new XAttribute("Type", WorksheetRelationshipType),
                    new XAttribute("Target", $"worksheets/sheet{index + 1}.xml")));
            }

            if (hasSharedStrings)
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", $"rId{relationshipId++}"),
                    new XAttribute("Type", SharedStringsRelationshipType),
                    new XAttribute("Target", "sharedStrings.xml")));
            }

            if (hasDateStyles)
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", $"rId{relationshipId++}"),
                    new XAttribute("Type", StylesRelationshipType),
                    new XAttribute("Target", "styles.xml")));
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), relationships);
        }

        internal static XDocument BuildWorksheet(WorksheetModel worksheet, StyleValue workbookDefaultStyle, Aspose.Cells_FOSS.Core.DateSystem dateSystem, SharedStringRepository sharedStrings, SaveOptions options, StylesheetSaveContext stylesheet)
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

            var hyperlinks = BuildHyperlinksElement(worksheet);
            if (hyperlinks != null)
            {
                worksheetElement.Add(hyperlinks);
            }

            worksheetElement.Add(BuildConditionalFormattingElements(worksheet, stylesheet));

            var dataValidations = BuildDataValidationsElement(worksheet);
            if (dataValidations != null)
            {
                worksheetElement.Add(dataValidations);
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

            var orderedIndexes = indexes.ToList();
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

        internal static string ToRangeReference(MergeRegion region)
        {
            var first = new CellAddress(region.FirstRow, region.FirstColumn).ToString();
            var last = new CellAddress(region.FirstRow + region.TotalRows - 1, region.FirstColumn + region.TotalColumns - 1).ToString();
            return string.Equals(first, last, StringComparison.Ordinal) ? first : first + ":" + last;
        }

        internal static bool TryParseMergeReference(string mergeReference, out MergeRegion region)
        {
            region = default(MergeRegion);
            if (string.IsNullOrWhiteSpace(mergeReference))
            {
                return false;
            }

            var parts = mergeReference.Split(':');
            if (parts.Length == 1)
            {
                CellAddress single;
                if (!TryParseCellReference(parts[0], out single))
                {
                    return false;
                }

                region = new MergeRegion(single.RowIndex, single.ColumnIndex, 1, 1);
                return true;
            }

            CellAddress first, last;
            if (parts.Length != 2
                || !TryParseCellReference(parts[0], out first)
                || !TryParseCellReference(parts[1], out last)
                || last.RowIndex < first.RowIndex
                || last.ColumnIndex < first.ColumnIndex)
            {
                return false;
            }

            region = new MergeRegion(first.RowIndex, first.ColumnIndex, last.RowIndex - first.RowIndex + 1, last.ColumnIndex - first.ColumnIndex + 1);
            return true;
        }

        internal static bool TryParseCellReference(string cellReference, out CellAddress address)
        {
            try
            {
                address = CellAddress.Parse(cellReference);
                return true;
            }
            catch (ArgumentException)
            {
                address = default(CellAddress);
                return false;
            }
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

        internal static XDocument BuildSharedStrings(SharedStringRepository sharedStrings)
        {
            var root = new XElement(MainNs + "sst",
                new XAttribute("count", sharedStrings.Values.Count),
                new XAttribute("uniqueCount", sharedStrings.Values.Count));

            foreach (var value in sharedStrings.Values)
            {
                root.Add(new XElement(MainNs + "si", CreateTextElement(value)));
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
        }

        internal static XElement CreateTextElement(string value)
        {
            var text = new XElement(MainNs + "t", value);
            if (NeedsPreserveWhitespace(value))
            {
                text.SetAttributeValue(XmlNs + "space", "preserve");
            }

            return text;
        }

        internal static bool NeedsPreserveWhitespace(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]) || value.Contains("\n") || value.Contains("\r") || value.Contains("\t");
        }

        internal static XDocument BuildMinimalStylesheet()
        {
            var stylesheet = new XElement(MainNs + "styleSheet",
                new XElement(MainNs + "fonts",
                    new XAttribute("count", 1),
                    new XElement(MainNs + "font",
                        new XElement(MainNs + "sz", new XAttribute("val", 11)),
                        new XElement(MainNs + "name", new XAttribute("val", "Calibri")))),
                new XElement(MainNs + "fills",
                    new XAttribute("count", 2),
                    new XElement(MainNs + "fill", new XElement(MainNs + "patternFill", new XAttribute("patternType", "none"))),
                    new XElement(MainNs + "fill", new XElement(MainNs + "patternFill", new XAttribute("patternType", "gray125")))),
                new XElement(MainNs + "borders",
                    new XAttribute("count", 1),
                    new XElement(MainNs + "border",
                        new XElement(MainNs + "left"),
                        new XElement(MainNs + "right"),
                        new XElement(MainNs + "top"),
                        new XElement(MainNs + "bottom"),
                        new XElement(MainNs + "diagonal"))),
                new XElement(MainNs + "cellStyleXfs",
                    new XAttribute("count", 1),
                    new XElement(MainNs + "xf",
                        new XAttribute("numFmtId", 0),
                        new XAttribute("fontId", 0),
                        new XAttribute("fillId", 0),
                        new XAttribute("borderId", 0))),
                new XElement(MainNs + "cellXfs",
                    new XAttribute("count", 2),
                    new XElement(MainNs + "xf",
                        new XAttribute("numFmtId", 0),
                        new XAttribute("fontId", 0),
                        new XAttribute("fillId", 0),
                        new XAttribute("borderId", 0),
                        new XAttribute("xfId", 0)),
                    new XElement(MainNs + "xf",
                        new XAttribute("numFmtId", 14),
                        new XAttribute("fontId", 0),
                        new XAttribute("fillId", 0),
                        new XAttribute("borderId", 0),
                        new XAttribute("xfId", 0),
                        new XAttribute("applyNumberFormat", 1))),
                new XElement(MainNs + "cellStyles",
                    new XAttribute("count", 1),
                    new XElement(MainNs + "cellStyle",
                        new XAttribute("name", "Normal"),
                        new XAttribute("xfId", 0),
                        new XAttribute("builtinId", 0))));

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), stylesheet);
        }

    }
}
