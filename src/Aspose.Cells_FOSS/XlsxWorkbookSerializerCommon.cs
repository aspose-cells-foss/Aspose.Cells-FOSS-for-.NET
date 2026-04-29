using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookStyles;
using static Aspose.Cells_FOSS.XlsxWorkbookDefinedNames;
using static Aspose.Cells_FOSS.XlsxWorkbookProperties;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookSerializerCommon
    {
        internal const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        internal const string SharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        internal const string StylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        internal const string ExternalLinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
        internal const string ExternalLinkContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
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

        internal static XDocument BuildContentTypes(WorkbookModel model, bool hasSharedStrings, bool hasDateStyles, bool hasCoreProperties, bool hasExtendedProperties, int totalTableCount, int totalDrawingCount, IReadOnlyList<string> imageExtensions, IReadOnlyList<string> chartPartNames, IReadOnlyList<string> chartContentTypes, IReadOnlyList<string> chartCompanionPartNames, IReadOnlyList<string> chartCompanionContentTypes, bool hasTheme = false, int totalCommentCount = 0)
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
                    new XAttribute("PartName", "/xl/worksheets/sheet" + (index + 1) + ".xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            }

            for (var tableNumber = 1; tableNumber <= totalTableCount; tableNumber++)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/xl/tables/table" + tableNumber + ".xml"),
                    new XAttribute("ContentType", XlsxWorkbookTables.TableContentType)));
            }

            for (var drawingNumber = 1; drawingNumber <= totalDrawingCount; drawingNumber++)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/xl/drawings/drawing" + drawingNumber + ".xml"),
                    new XAttribute("ContentType", XlsxWorkbookPictures.DrawingContentType)));
            }

            for (var c = 0; c < chartPartNames.Count; c++)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", chartPartNames[c]),
                    new XAttribute("ContentType", chartContentTypes[c])));
            }

            for (var c = 0; c < chartCompanionPartNames.Count; c++)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", chartCompanionPartNames[c]),
                    new XAttribute("ContentType", chartCompanionContentTypes[c])));
            }

            for (var extIndex = 0; extIndex < imageExtensions.Count; extIndex++)
            {
                var ext = imageExtensions[extIndex];
                root.Add(new XElement(ContentTypeNs + "Default",
                    new XAttribute("Extension", ext),
                    new XAttribute("ContentType", Picture.ContentTypeFromExtension(ext))));
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

            if (hasTheme)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/xl/theme/theme1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml")));
            }

            for (var n = 1; n <= totalCommentCount; n++)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/xl/comments" + n.ToString(CultureInfo.InvariantCulture) + ".xml"),
                    new XAttribute("ContentType", XlsxWorkbookComments.CommentsContentType)));
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/xl/drawings/vmlDrawing" + n.ToString(CultureInfo.InvariantCulture) + ".vml"),
                    new XAttribute("ContentType", XlsxWorkbookComments.VmlDrawingContentType)));
            }

            for (var i = 0; i < model.ExternalLinks.Count; i++)
            {
                root.Add(new XElement(ContentTypeNs + "Override",
                    new XAttribute("PartName", "/xl/externalLinks/externalLink" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml"),
                    new XAttribute("ContentType", ExternalLinkContentType)));
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

        internal static XDocument BuildWorkbook(WorkbookModel model, int externalLinkBaseRId)
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

            if (model.ExternalLinks.Count > 0)
            {
                var externalReferences = new XElement(MainNs + "externalReferences");
                for (var i = 0; i < model.ExternalLinks.Count; i++)
                {
                    externalReferences.Add(new XElement(MainNs + "externalReference",
                        new XAttribute(RelationshipNs + "id", "rId" + (externalLinkBaseRId + i).ToString(CultureInfo.InvariantCulture))));
                }

                workbook.Add(externalReferences);
            }

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

            if (!string.IsNullOrEmpty(model.RawThemeXml))
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", $"rId{relationshipId++}"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
                    new XAttribute("Target", "theme/theme1.xml")));
            }

            for (var i = 0; i < model.ExternalLinks.Count; i++)
            {
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + relationshipId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", ExternalLinkRelationshipType),
                    new XAttribute("Target", "externalLinks/externalLink" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml")));
                relationshipId++;
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), relationships);
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
