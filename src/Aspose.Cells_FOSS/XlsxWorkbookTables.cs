using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;
using static Aspose.Cells_FOSS.XlsxWorkbookHyperlinks;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookTables
    {
        internal const string TableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
        internal const string TableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

        internal static XDocument BuildTableDocument(ListObjectModel model, int tableId)
        {
            var startAddress = new CellAddress(model.StartRow, model.StartColumn).ToString();
            var endAddress = new CellAddress(model.EndRow, model.EndColumn).ToString();
            var tableRef = startAddress + ":" + endAddress;

            var tableElement = new XElement(MainNs + "table",
                new XAttribute(XNamespace.Xmlns + "r", RelationshipNs),
                new XAttribute("id", tableId.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("name", model.Name),
                new XAttribute("displayName", model.DisplayName),
                new XAttribute("ref", tableRef));

            if (!model.ShowHeaderRow)
            {
                tableElement.SetAttributeValue("headerRowCount", "0");
            }

            if (model.ShowTotals)
            {
                tableElement.SetAttributeValue("totalsRowCount", "1");
            }
            else
            {
                tableElement.SetAttributeValue("totalsRowShown", "0");
            }

            if (!string.IsNullOrEmpty(model.Comment))
            {
                tableElement.SetAttributeValue("comment", model.Comment);
            }

            if (model.HasAutoFilter && model.ShowHeaderRow)
            {
                var autoFilterRef = BuildAutoFilterRef(model);
                tableElement.Add(new XElement(MainNs + "autoFilter", new XAttribute("ref", autoFilterRef)));
            }

            var tableColumns = new XElement(MainNs + "tableColumns",
                new XAttribute("count", model.Columns.Count));

            for (var i = 0; i < model.Columns.Count; i++)
            {
                tableColumns.Add(BuildTableColumnElement(model.Columns[i]));
            }

            tableElement.Add(tableColumns);

            var tableStyleInfo = BuildTableStyleInfoElement(model);
            if (tableStyleInfo != null)
            {
                tableElement.Add(tableStyleInfo);
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), tableElement);
        }

        private static string BuildAutoFilterRef(ListObjectModel model)
        {
            var dataEndRow = model.ShowTotals ? model.EndRow - 1 : model.EndRow;
            var startAddress = new CellAddress(model.StartRow, model.StartColumn).ToString();
            var endAddress = new CellAddress(dataEndRow, model.EndColumn).ToString();
            return startAddress + ":" + endAddress;
        }

        private static XElement BuildTableColumnElement(ListColumnModel column)
        {
            var element = new XElement(MainNs + "tableColumn",
                new XAttribute("id", column.Id.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("name", column.Name));

            var function = column.TotalsRowFunction;
            if (!string.IsNullOrEmpty(function) && !string.Equals(function, "none", StringComparison.Ordinal))
            {
                element.SetAttributeValue("totalsRowFunction", function);
            }

            if (string.Equals(function, "none", StringComparison.Ordinal) || string.IsNullOrEmpty(function))
            {
                if (!string.IsNullOrEmpty(column.TotalsRowLabel))
                {
                    element.SetAttributeValue("totalsRowLabel", column.TotalsRowLabel);
                }
            }

            if (string.Equals(function, "custom", StringComparison.Ordinal) && !string.IsNullOrEmpty(column.TotalsRowFormula))
            {
                element.SetAttributeValue("totalsRowFormula", column.TotalsRowFormula);
            }

            return element;
        }

        private static XElement BuildTableStyleInfoElement(ListObjectModel model)
        {
            if (string.IsNullOrEmpty(model.TableStyleName))
            {
                return null;
            }

            var element = new XElement(MainNs + "tableStyleInfo",
                new XAttribute("name", model.TableStyleName),
                new XAttribute("showFirstColumn", model.ShowFirstColumn ? "1" : "0"),
                new XAttribute("showLastColumn", model.ShowLastColumn ? "1" : "0"),
                new XAttribute("showRowStripes", model.ShowRowStripes ? "1" : "0"),
                new XAttribute("showColumnStripes", model.ShowColumnStripes ? "1" : "0"));

            return element;
        }

        internal static XElement BuildTablePartsElement(WorksheetModel worksheet, int tableRIdStart)
        {
            if (worksheet.ListObjects.Count == 0)
            {
                return null;
            }

            var tablePartsElement = new XElement(MainNs + "tableParts",
                new XAttribute("count", worksheet.ListObjects.Count));

            for (var i = 0; i < worksheet.ListObjects.Count; i++)
            {
                var rId = "rId" + (tableRIdStart + i).ToString(CultureInfo.InvariantCulture);
                tablePartsElement.Add(new XElement(MainNs + "tablePart",
                    new XAttribute(RelationshipNs + "id", rId)));
            }

            return tablePartsElement;
        }

        internal static XDocument BuildWorksheetRelationshipsDocument(WorksheetModel worksheet, int tableFileOffset, int drawingNumber, int commentFileNumber)
        {
            var relationships = new XElement(PackageRelationshipNs + "Relationships");
            var rId = 1;

            var orderedHyperlinks = GetOrderedHyperlinks(worksheet.Hyperlinks);
            for (var i = 0; i < orderedHyperlinks.Count; i++)
            {
                var hyperlink = orderedHyperlinks[i];
                if (string.IsNullOrEmpty(hyperlink.Address))
                {
                    continue;
                }

                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + rId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", HyperlinkRelationshipType),
                    new XAttribute("Target", hyperlink.Address),
                    new XAttribute("TargetMode", "External")));
                rId++;
            }

            for (var t = 0; t < worksheet.ListObjects.Count; t++)
            {
                var globalTableNumber = tableFileOffset + t + 1;
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + rId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", TableRelationshipType),
                    new XAttribute("Target", "../tables/table" + globalTableNumber.ToString(CultureInfo.InvariantCulture) + ".xml")));
                rId++;
            }

            if (drawingNumber > 0)
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + rId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", XlsxWorkbookPictures.DrawingRelationshipType),
                    new XAttribute("Target", "../drawings/drawing" + drawingNumber.ToString(CultureInfo.InvariantCulture) + ".xml")));
                rId++;
            }

            if (commentFileNumber > 0)
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + rId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", XlsxWorkbookComments.CommentsRelationshipType),
                    new XAttribute("Target", "../comments" + commentFileNumber.ToString(CultureInfo.InvariantCulture) + ".xml")));
                rId++;
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + rId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", XlsxWorkbookComments.VmlDrawingRelationshipType),
                    new XAttribute("Target", "../drawings/vmlDrawing" + commentFileNumber.ToString(CultureInfo.InvariantCulture) + ".vml")));
                rId++;
            }

            if (!relationships.HasElements)
            {
                return null;
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), relationships);
        }

        internal static IReadOnlyDictionary<string, string> LoadWorksheetTableTargets(ZipArchive archive, string worksheetUri)
        {
            var relsUri = GetWorksheetRelsUri(worksheetUri);
            var entry = GetEntry(archive, relsUri);
            if (entry == null)
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            var document = LoadDocument(entry);
            var targets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var rel in document.Root != null ? document.Root.Elements(PackageRelationshipNs + "Relationship") : new XElement[0])
            {
                var id = (string)rel.Attribute("Id");
                var type = (string)rel.Attribute("Type");
                var target = (string)rel.Attribute("Target");
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target))
                {
                    continue;
                }

                if (!string.Equals(type, TableRelationshipType, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                targets[id] = ResolvePartUri(worksheetUri, target);
            }

            return targets;
        }

        internal static void LoadTables(WorksheetModel worksheetModel, XElement worksheetRoot, ZipArchive archive, string worksheetUri, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            worksheetModel.ListObjects.Clear();

            var tablePartsElement = worksheetRoot.Element(MainNs + "tableParts");
            if (tablePartsElement == null)
            {
                return;
            }

            var tableTargets = LoadWorksheetTableTargets(archive, worksheetUri);

            foreach (var tablePartElement in tablePartsElement.Elements(MainNs + "tablePart"))
            {
                var rId = (string)tablePartElement.Attribute(RelationshipNs + "id");
                if (string.IsNullOrEmpty(rId))
                {
                    continue;
                }

                string tableUri;
                if (!tableTargets.TryGetValue(rId, out tableUri))
                {
                    AddIssue(diagnostics, options, new LoadIssue("TBL-L002", DiagnosticSeverity.LossyRecoverable, "Table relationship '" + rId + "' could not be resolved and the table was dropped.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                    });
                    continue;
                }

                var tableEntry = GetEntry(archive, tableUri);
                if (tableEntry == null)
                {
                    AddIssue(diagnostics, options, new LoadIssue("TBL-R001", DiagnosticSeverity.Recoverable, "Table part '" + tableUri + "' was not found and the table was skipped.", repairApplied: true)
                    {
                        SheetName = sheetName,
                    });
                    continue;
                }

                var tableDocument = LoadDocument(tableEntry);
                var tableModel = LoadTableModel(tableDocument, diagnostics, options, sheetName, worksheetModel);
                if (tableModel != null)
                {
                    worksheetModel.ListObjects.Add(tableModel);
                }
            }
        }

        private static ListObjectModel LoadTableModel(XDocument tableDocument, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, WorksheetModel worksheetModel)
        {
            var root = tableDocument.Root;
            if (root == null)
            {
                return null;
            }

            var tableRef = (string)root.Attribute("ref");
            if (string.IsNullOrEmpty(tableRef))
            {
                AddIssue(diagnostics, options, new LoadIssue("TBL-L001", DiagnosticSeverity.LossyRecoverable, "Table ref attribute is missing and the table was dropped.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                });
                return null;
            }

            int startRow, startColumn, endRow, endColumn;
            if (!TryParseTableRef(tableRef, out startRow, out startColumn, out endRow, out endColumn))
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("Table ref '" + tableRef + "' is invalid.");
                }

                AddIssue(diagnostics, options, new LoadIssue("TBL-L001", DiagnosticSeverity.LossyRecoverable, "Table ref '" + tableRef + "' is invalid and the table was dropped.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                });
                return null;
            }

            var model = new ListObjectModel
            {
                StartRow = startRow,
                StartColumn = startColumn,
                EndRow = endRow,
                EndColumn = endColumn,
            };

            model.DisplayName = (string)root.Attribute("displayName") ?? string.Empty;
            model.Name = (string)root.Attribute("name") ?? model.DisplayName;
            model.Comment = (string)root.Attribute("comment") ?? string.Empty;

            if (string.IsNullOrEmpty(model.DisplayName))
            {
                model.DisplayName = model.Name;
            }

            var headerRowCount = ParseIntAttribute(root.Attribute("headerRowCount"));
            model.ShowHeaderRow = !headerRowCount.HasValue || headerRowCount.Value != 0;

            var totalsRowCount = ParseIntAttribute(root.Attribute("totalsRowCount"));
            model.ShowTotals = totalsRowCount.HasValue && totalsRowCount.Value > 0;

            var autoFilterElement = root.Element(MainNs + "autoFilter");
            model.HasAutoFilter = autoFilterElement != null;

            LoadTableColumns(model, root, diagnostics, options, sheetName);
            LoadTableStyleInfo(model, root);

            return model;
        }

        private static void LoadTableColumns(ListObjectModel model, XElement root, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var tableColumnsElement = root.Element(MainNs + "tableColumns");
            if (tableColumnsElement == null)
            {
                var expectedCount = model.EndColumn - model.StartColumn + 1;
                for (var c = 0; c < expectedCount; c++)
                {
                    model.Columns.Add(new ListColumnModel(c + 1, "Column" + (c + 1).ToString(CultureInfo.InvariantCulture)));
                }

                AddIssue(diagnostics, options, new LoadIssue("TBL-R004", DiagnosticSeverity.Recoverable, "Table columns element was missing; default column names were synthesized.", repairApplied: true)
                {
                    SheetName = sheetName,
                });
                return;
            }

            foreach (var columnElement in tableColumnsElement.Elements(MainNs + "tableColumn"))
            {
                var idAttr = ParseIntAttribute(columnElement.Attribute("id"));
                var name = (string)columnElement.Attribute("name") ?? string.Empty;
                var columnId = idAttr.HasValue && idAttr.Value > 0 ? idAttr.Value : model.Columns.Count + 1;

                var columnModel = new ListColumnModel(columnId, name)
                {
                    TotalsRowFunction = (string)columnElement.Attribute("totalsRowFunction") ?? "none",
                    TotalsRowLabel = (string)columnElement.Attribute("totalsRowLabel") ?? string.Empty,
                    TotalsRowFormula = (string)columnElement.Attribute("totalsRowFormula") ?? string.Empty,
                };

                model.Columns.Add(columnModel);
            }
        }

        private static void LoadTableStyleInfo(ListObjectModel model, XElement root)
        {
            var styleInfoElement = root.Element(MainNs + "tableStyleInfo");
            if (styleInfoElement == null)
            {
                return;
            }

            model.TableStyleName = (string)styleInfoElement.Attribute("name") ?? string.Empty;
            model.ShowFirstColumn = ParseBoolAttribute(styleInfoElement.Attribute("showFirstColumn"));
            model.ShowLastColumn = ParseBoolAttribute(styleInfoElement.Attribute("showLastColumn"));
            model.ShowRowStripes = ParseBoolAttribute(styleInfoElement.Attribute("showRowStripes"));
            model.ShowColumnStripes = ParseBoolAttribute(styleInfoElement.Attribute("showColumnStripes"));
        }

        private static bool TryParseTableRef(string tableRef, out int startRow, out int startColumn, out int endRow, out int endColumn)
        {
            startRow = 0;
            startColumn = 0;
            endRow = 0;
            endColumn = 0;

            if (string.IsNullOrWhiteSpace(tableRef))
            {
                return false;
            }

            var parts = tableRef.Split(':');
            if (parts.Length != 2)
            {
                return false;
            }

            CellAddress first, last;
            if (!TryParseCellReference(parts[0], out first) || !TryParseCellReference(parts[1], out last))
            {
                return false;
            }

            if (last.RowIndex < first.RowIndex || last.ColumnIndex < first.ColumnIndex)
            {
                return false;
            }

            startRow = first.RowIndex;
            startColumn = first.ColumnIndex;
            endRow = last.RowIndex;
            endColumn = last.ColumnIndex;
            return true;
        }

        private static string GetWorksheetRelsUri(string worksheetUri)
        {
            var normalized = worksheetUri.TrimStart('/');
            var slashIndex = normalized.LastIndexOf('/');
            var directory = slashIndex >= 0 ? normalized.Substring(0, slashIndex + 1) : string.Empty;
            var fileName = slashIndex >= 0 ? normalized.Substring(slashIndex + 1) : normalized;
            return "/" + directory + "_rels/" + fileName + ".rels";
        }
    }
}
