using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookConditionalFormatting;
using static Aspose.Cells_FOSS.XlsxWorkbookHyperlinks;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;
using static Aspose.Cells_FOSS.XlsxWorkbookStyles;
using static Aspose.Cells_FOSS.XlsxWorkbookPageSetup;
using static Aspose.Cells_FOSS.XlsxWorkbookValidations;
using static Aspose.Cells_FOSS.XlsxWorkbookWorksheetProtection;
using static Aspose.Cells_FOSS.XlsxWorkbookAutoFilter;
using static Aspose.Cells_FOSS.XlsxWorkbookWorksheetViews;
using static Aspose.Cells_FOSS.XlsxWorkbookTables;
using static Aspose.Cells_FOSS.XlsxWorkbookPicturesLoader;
using static Aspose.Cells_FOSS.XlsxWorkbookCommentsLoader;
using static Aspose.Cells_FOSS.XlsxWorkbookDefinedNames;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookWorksheetLoader
    {
        internal static WorksheetModel LoadWorksheet(
            string sheetName,
            XDocument worksheetDocument,
            IReadOnlyDictionary<string, string> worksheetHyperlinkTargets,
            Aspose.Cells_FOSS.Core.DateSystem dateSystem,
            IReadOnlyList<string> sharedStrings,
            StylesheetLoadContext stylesheet,
            LoadDiagnostics diagnostics,
            LoadOptions options,
            WorksheetDefinedNamesState definedNamesState,
            ZipArchive archive,
            string worksheetUri)
        {
            var worksheetModel = new WorksheetModel(sheetName);
            var worksheetRoot = worksheetDocument.Root;
            if (worksheetRoot == null)
            {
                return worksheetModel;
            }

            LoadWorksheetViewSettings(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
            LoadWorksheetProtection(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
            LoadAutoFilter(worksheetModel, worksheetRoot, stylesheet, diagnostics, options, sheetName);
            LoadColumns(worksheetModel, worksheetRoot, stylesheet, diagnostics, options, sheetName);
            LoadWorksheetPageSetup(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
            LoadHyperlinks(worksheetModel, worksheetRoot, worksheetHyperlinkTargets, diagnostics, options, sheetName);
            LoadConditionalFormattings(worksheetModel, worksheetRoot, stylesheet, diagnostics, options, sheetName);
            LoadValidations(worksheetModel, worksheetRoot, diagnostics, options, sheetName);
            LoadTables(worksheetModel, worksheetRoot, archive, worksheetUri, diagnostics, options, sheetName);
            LoadPictures(worksheetModel, worksheetRoot, archive, worksheetUri, diagnostics, options, sheetName);
            LoadComments(worksheetModel, archive, worksheetUri, diagnostics, options, sheetName);
            ApplyWorksheetDefinedNames(worksheetModel, definedNamesState);

            var sheetData = worksheetRoot.Element(MainNs + "sheetData");
            if (sheetData == null)
            {
                AddIssue(diagnostics, options, new LoadIssue("ACF-WS-001", DiagnosticSeverity.Recoverable, "Worksheet sheetData is missing; an empty sheet was synthesized.", repairApplied: true)
                {
                    SheetName = sheetName,
                });
                return worksheetModel;
            }

            var rowElements = new List<XElement>(sheetData.Elements(MainNs + "row"));
            var seenRows = new HashSet<int>();
            var seenCells = new HashSet<CellAddress>();
            var previousRowIndex = -1;
            var rowsOutOfOrderReported = false;

            foreach (var rowElement in rowElements)
            {
                int rowIndex;
                if (!TryResolveRowIndex(rowElement, diagnostics, options, sheetName, out rowIndex))
                {
                    continue;
                }

                if (!seenRows.Add(rowIndex))
                {
                    throw new InvalidFileFormatException($"Duplicate row index '{rowIndex + 1}' was found in worksheet '{sheetName}'.");
                }

                if (previousRowIndex > rowIndex && !rowsOutOfOrderReported)
                {
                    AddIssue(diagnostics, options, new LoadIssue("WS-R002", DiagnosticSeverity.Recoverable, "Worksheet rows were out of order and were normalized during load.", repairApplied: true)
                    {
                        SheetName = sheetName,
                    });
                    rowsOutOfOrderReported = true;
                }

                previousRowIndex = rowIndex;
                ApplyRowMetadata(worksheetModel, rowElement, rowIndex, stylesheet, diagnostics, options, sheetName);

                var previousColumnIndex = -1;
                var cellsOutOfOrderReported = false;
                foreach (var cellElement in rowElement.Elements(MainNs + "c"))
                {
                    CellAddress address;
                    CellRecord record;
                    if (!TryReadCellRecord(cellElement, rowIndex, dateSystem, sharedStrings, stylesheet, diagnostics, options, sheetName, out address, out record))
                    {
                        continue;
                    }

                    if (previousColumnIndex > address.ColumnIndex && !cellsOutOfOrderReported)
                    {
                        AddIssue(diagnostics, options, new LoadIssue("WS-R003", DiagnosticSeverity.Recoverable, "Worksheet cells were out of order within a row and were normalized during load.", repairApplied: true)
                        {
                            SheetName = sheetName,
                            RowIndex = rowIndex,
                        });
                        cellsOutOfOrderReported = true;
                    }

                    previousColumnIndex = address.ColumnIndex;
                    if (!seenCells.Add(address))
                    {
                        throw new InvalidFileFormatException($"Duplicate cell reference '{address}' was found in worksheet '{sheetName}'.");
                    }

                    if (ShouldPersistCell(stylesheet.DefaultCellStyle, record))
                    {
                        worksheetModel.Cells[address] = record;
                    }
                }
            }

            LoadMergeRegions(worksheetModel, worksheetRoot, diagnostics, options, sheetName);

            if (worksheetRoot.Element(MainNs + "dimension") == null && (worksheetModel.Cells.Count > 0 || worksheetModel.MergeRegions.Count > 0))
            {
                AddIssue(diagnostics, options, new LoadIssue("WS-R001", DiagnosticSeverity.Recoverable, "Worksheet dimension was missing and was recalculated during load.", repairApplied: true)
                {
                    SheetName = sheetName,
                });
            }

            return worksheetModel;
        }

        private static void LoadColumns(WorksheetModel worksheetModel, XElement worksheetRoot, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var columns = new List<ColumnRangeModel>();
            foreach (var columnElement in worksheetRoot.Element(MainNs + "cols")?.Elements(MainNs + "col") ?? new XElement[0])
            {
                var min = ParseIntAttribute(columnElement.Attribute("min"));
                var max = ParseIntAttribute(columnElement.Attribute("max"));
                if (!min.HasValue || !max.HasValue || min.Value <= 0 || max.Value <= 0 || min.Value > max.Value)
                {
                    throw new InvalidFileFormatException($"Worksheet column metadata in '{sheetName}' contains an invalid min/max span.");
                }

                var styleIndex = ParseIntAttribute(columnElement.Attribute("style"));
                if (styleIndex.HasValue && (styleIndex.Value < 0 || styleIndex.Value >= stylesheet.CellFormats.Count))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException($"The column style index '{styleIndex.Value}' is invalid.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("COL-L001", DiagnosticSeverity.Warning, $"Column style index '{styleIndex.Value}' is invalid and was dropped.", dataLossRisk: true)
                    {
                        SheetName = sheetName,
                    });
                    styleIndex = null;
                }

                columns.Add(new ColumnRangeModel
                {
                    MinColumnIndex = min.Value - 1,
                    MaxColumnIndex = max.Value - 1,
                    Width = ParseDoubleAttribute(columnElement.Attribute("width")),
                    Hidden = ParseBoolAttribute(columnElement.Attribute("hidden")),
                    StyleIndex = styleIndex,
                });
            }

            worksheetModel.Columns.AddRange(NormalizeLoadedColumns(columns, diagnostics, options, sheetName));
        }

        private static List<ColumnRangeModel> NormalizeLoadedColumns(IReadOnlyList<ColumnRangeModel> columns, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var ordered = new List<ColumnRangeModel>(columns.Count);
            for (var index = 0; index < columns.Count; index++)
            {
                ordered.Add(columns[index]);
            }

            ordered.Sort(CompareColumnRangesByBounds);
            if (ordered.Count == 0)
            {
                return ordered;
            }

            var normalized = new List<ColumnRangeModel> { ordered[0] };
            var mergeReported = false;
            for (var index = 1; index < ordered.Count; index++)
            {
                var current = ordered[index];
                var previous = normalized[normalized.Count - 1];
                if (current.MinColumnIndex <= previous.MaxColumnIndex + 1 && ColumnRangesCompatible(previous, current))
                {
                    if ((current.MinColumnIndex <= previous.MaxColumnIndex || current.MaxColumnIndex <= previous.MaxColumnIndex) && !mergeReported)
                    {
                        AddIssue(diagnostics, options, new LoadIssue("COL-R001", DiagnosticSeverity.Recoverable, "Overlapping compatible column metadata was normalized during load.", repairApplied: true)
                        {
                            SheetName = sheetName,
                        });
                        mergeReported = true;
                    }

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

        private static bool ColumnRangesCompatible(ColumnRangeModel left, ColumnRangeModel right)
        {
            return Nullable.Equals(left.Width, right.Width)
                && left.Hidden == right.Hidden
                && left.StyleIndex == right.StyleIndex;
        }

        private static bool TryResolveRowIndex(XElement rowElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, out int rowIndex)
        {
            var rowIndexAttribute = ParseIntAttribute(rowElement.Attribute("r"));
            if (rowIndexAttribute.HasValue)
            {
                if (rowIndexAttribute.Value <= 0)
                {
                    throw new InvalidFileFormatException($"Worksheet row index '{rowIndexAttribute.Value}' is invalid.");
                }

                rowIndex = rowIndexAttribute.Value - 1;
                return true;
            }

            foreach (var cellElement in rowElement.Elements(MainNs + "c"))
            {
                var cellReference = (string)cellElement.Attribute("r");
                CellAddress address;
                if (TryParseCellReference(cellReference ?? string.Empty, out address))
                {
                    rowIndex = address.RowIndex;
                    AddIssue(diagnostics, options, new LoadIssue("WS-R004", DiagnosticSeverity.Recoverable, "A worksheet row index was missing and was inferred from contained cells.", repairApplied: true)
                    {
                        SheetName = sheetName,
                        RowIndex = rowIndex,
                    });
                    return true;
                }
            }

            rowIndex = -1;
            AddIssue(diagnostics, options, new LoadIssue("ROW-F001", DiagnosticSeverity.Warning, "A worksheet row without an index and without parseable cells was skipped.")
            {
                SheetName = sheetName,
            });
            return false;
        }

        private static void ApplyRowMetadata(WorksheetModel worksheetModel, XElement rowElement, int rowIndex, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var styleIndex = ParseIntAttribute(rowElement.Attribute("s"));
            if (styleIndex.HasValue && (styleIndex.Value < 0 || styleIndex.Value >= stylesheet.CellFormats.Count))
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException($"The row style index '{styleIndex.Value}' is invalid.");
                }

                AddIssue(diagnostics, options, new LoadIssue("ROW-L001", DiagnosticSeverity.Warning, $"Row style index '{styleIndex.Value}' is invalid and was dropped.", dataLossRisk: true)
                {
                    SheetName = sheetName,
                    RowIndex = rowIndex,
                });
                styleIndex = null;
            }

            var rowModel = new RowModel
            {
                Height = ParseDoubleAttribute(rowElement.Attribute("ht")),
                Hidden = ParseBoolAttribute(rowElement.Attribute("hidden")),
                StyleIndex = styleIndex,
            };

            if (rowModel.Height.HasValue && !ParseBoolAttribute(rowElement.Attribute("customHeight")))
            {
                AddIssue(diagnostics, options, new LoadIssue("ROW-R002", DiagnosticSeverity.Recoverable, "Row height metadata was missing customHeight and was normalized during load.", repairApplied: true)
                {
                    SheetName = sheetName,
                    RowIndex = rowIndex,
                });
            }

            if (rowModel.Height.HasValue || rowModel.Hidden || rowModel.StyleIndex.HasValue)
            {
                worksheetModel.Rows[rowIndex] = rowModel;
            }
        }

        private static bool TryReadCellRecord(
            XElement cellElement,
            int rowIndex,
            Aspose.Cells_FOSS.Core.DateSystem dateSystem,
            IReadOnlyList<string> sharedStrings,
            StylesheetLoadContext stylesheet,
            LoadDiagnostics diagnostics,
            LoadOptions options,
            string sheetName,
            out CellAddress address,
            out CellRecord record)
        {
            record = new CellRecord
            {
                Style = stylesheet.DefaultCellStyle.Clone(),
                IsExplicitlyStored = true,
            };

            var cellReference = (string)cellElement.Attribute("r");
            if (string.IsNullOrWhiteSpace(cellReference))
            {
                AddIssue(diagnostics, options, new LoadIssue("CELL-F001", DiagnosticSeverity.Warning, "A cell without a reference was skipped.")
                {
                    SheetName = sheetName,
                    RowIndex = rowIndex,
                });
                address = default(CellAddress);
                return false;
            }

            var resolvedCellReference = cellReference;

            try
            {
                address = CellAddress.Parse(resolvedCellReference);
            }
            catch (ArgumentException)
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException($"The cell reference '{resolvedCellReference}' is invalid.");
                }

                AddIssue(diagnostics, options, new LoadIssue("CELL-F001", DiagnosticSeverity.Warning, $"Cell reference '{resolvedCellReference}' is invalid and was skipped.")
                {
                    SheetName = sheetName,
                    CellRef = resolvedCellReference,
                    RowIndex = rowIndex,
                });
                address = default(CellAddress);
                return false;
            }

            var styleIndex = ParseIntAttribute(cellElement.Attribute("s"));
            var isDateStyle = styleIndex.HasValue && stylesheet.DateStyleIndexes.Contains(styleIndex.Value);
            if (styleIndex.HasValue)
            {
                if (styleIndex.Value >= 0 && styleIndex.Value < stylesheet.CellFormats.Count)
                {
                    record.Style = stylesheet.CellFormats[styleIndex.Value].Clone();
                }
                else if (options.StrictMode)
                {
                    throw new InvalidFileFormatException($"The style index '{styleIndex.Value}' is invalid.");
                }
                else
                {
                    AddIssue(diagnostics, options, new LoadIssue("STYLE-F001", DiagnosticSeverity.Warning, $"Cell style index '{styleIndex.Value}' is invalid and style 0 was used instead.")
                    {
                        SheetName = sheetName,
                        CellRef = cellReference,
                        RowIndex = rowIndex,
                    });
                }
            }

            var formulaText = NormalizeFormula((string)cellElement.Element(MainNs + "f"));
            if (!string.IsNullOrEmpty(formulaText))
            {
                record.Formula = formulaText;
                record.Kind = CellValueKind.Formula;
            }

            var cellType = (string)cellElement.Attribute("t");
            var valueElement = cellElement.Element(MainNs + "v");
            object value;
            CellValueKind kind;
            if (TryReadCellValue(cellElement, cellType, valueElement?.Value, isDateStyle, dateSystem, sharedStrings, diagnostics, options, sheetName, resolvedCellReference, out value, out kind))
            {
                record.Value = value;
                if (string.IsNullOrEmpty(record.Formula) || kind == CellValueKind.Error)
                {
                    record.Kind = kind;
                }
            }

            return true;
        }

        private static void LoadMergeRegions(WorksheetModel worksheetModel, XElement worksheetRoot, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            foreach (var mergeElement in worksheetRoot.Element(MainNs + "mergeCells")?.Elements(MainNs + "mergeCell") ?? new XElement[0])
            {
                var mergeReference = (string)mergeElement.Attribute("ref");
                MergeRegion region;
                if (!TryParseMergeReference(mergeReference ?? string.Empty, out region))
                {
                    throw new InvalidFileFormatException($"The merge reference '{mergeReference}' is invalid.");
                }

                if (ContainsOverlappingMergeRegion(worksheetModel.MergeRegions, region))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException($"The merge reference '{mergeReference}' overlaps an existing merged range.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("MRG-L001", DiagnosticSeverity.LossyRecoverable, $"Overlapping merge range '{mergeReference}' was dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                    });
                    continue;
                }

                worksheetModel.MergeRegions.Add(region);
            }

            worksheetModel.MergeRegions.Sort(delegate(MergeRegion left, MergeRegion right)
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
                return rowCountComparison != 0 ? rowCountComparison : left.TotalColumns.CompareTo(right.TotalColumns);
            });
        }

        private static bool ContainsOverlappingMergeRegion(IReadOnlyList<MergeRegion> mergeRegions, MergeRegion candidate)
        {
            for (var index = 0; index < mergeRegions.Count; index++)
            {
                if (MergeRegionsOverlap(mergeRegions[index], candidate))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool MergeRegionsOverlap(MergeRegion left, MergeRegion right)
        {
            var leftLastRow = left.FirstRow + left.TotalRows - 1;
            var leftLastColumn = left.FirstColumn + left.TotalColumns - 1;
            var rightLastRow = right.FirstRow + right.TotalRows - 1;
            var rightLastColumn = right.FirstColumn + right.TotalColumns - 1;

            return left.FirstRow <= rightLastRow
                && right.FirstRow <= leftLastRow
                && left.FirstColumn <= rightLastColumn
                && right.FirstColumn <= leftLastColumn;
        }

        private static bool TryReadCellValue(
            XElement cellElement,
            string cellType,
            string rawValue,
            bool isDateStyle,
            Aspose.Cells_FOSS.Core.DateSystem dateSystem,
            IReadOnlyList<string> sharedStrings,
            LoadDiagnostics diagnostics,
            LoadOptions options,
            string sheetName,
            string cellReference,
            out object value,
            out CellValueKind kind)
        {
            value = null;
            kind = CellValueKind.Blank;

            if (cellType == "inlineStr")
            {
                value = ReadInlineString(cellElement.Element(MainNs + "is"));
                kind = CellValueKind.String;
                return true;
            }

            if (cellType == "s")
            {
                int sharedStringIndex;
                if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out sharedStringIndex) && sharedStringIndex >= 0 && sharedStringIndex < sharedStrings.Count)
                {
                    value = sharedStrings[sharedStringIndex];
                    kind = CellValueKind.String;
                    return true;
                }

                AddIssue(diagnostics, options, new LoadIssue("SST-L001", DiagnosticSeverity.LossyRecoverable, "The cell points to an invalid shared string index.", dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = cellReference,
                });

                value = string.Empty;
                kind = CellValueKind.String;
                return true;
            }

            if (cellType == "b")
            {
                value = rawValue == "1" || string.Equals(rawValue, "true", StringComparison.OrdinalIgnoreCase);
                kind = CellValueKind.Boolean;
                return true;
            }

            if (cellType == "d")
            {
                DateTime dateTime;
                if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out dateTime))
                {
                    value = dateTime;
                    kind = CellValueKind.DateTime;
                    return true;
                }

                return false;
            }

            if (cellType == "str")
            {
                value = rawValue ?? string.Empty;
                kind = CellValueKind.String;
                return true;
            }

            if (cellType == "e")
            {
                value = rawValue ?? string.Empty;
                kind = CellValueKind.Error;
                return true;
            }

            if (string.IsNullOrEmpty(rawValue))
            {
                return false;
            }

            var resolvedRawValue = rawValue;

            if (isDateStyle)
            {
                double serial;
                if (double.TryParse(resolvedRawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out serial))
                {
                    value = DateSerialConverter.FromSerial(serial, dateSystem);
                    kind = CellValueKind.DateTime;
                    return true;
                }

                AddIssue(diagnostics, options, new LoadIssue("CELL-R002", DiagnosticSeverity.Recoverable, "A formula or numeric cell contained an invalid cached date serial and the cached value was cleared.", repairApplied: true)
                {
                    SheetName = sheetName,
                    CellRef = cellReference,
                });
                return false;
            }

            object numberValue;
            if (TryParseNumber(resolvedRawValue, out numberValue))
            {
                value = numberValue;
                kind = CellValueKind.Number;
                return true;
            }

            return false;
        }
    }
}
