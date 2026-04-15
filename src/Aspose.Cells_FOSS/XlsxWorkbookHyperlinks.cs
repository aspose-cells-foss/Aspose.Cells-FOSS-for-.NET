using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookHyperlinks
    {
        private const string HyperlinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        internal static XElement BuildHyperlinksElement(WorksheetModel worksheet)
        {
            var orderedHyperlinks = GetOrderedHyperlinks(worksheet.Hyperlinks);
            if (orderedHyperlinks.Count == 0)
            {
                return null;
            }

            var relationshipId = 1;
            var hyperlinksElement = new XElement(MainNs + "hyperlinks");
            for (var index = 0; index < orderedHyperlinks.Count; index++)
            {
                var hyperlink = orderedHyperlinks[index];
                if (!ShouldPersistHyperlink(hyperlink))
                {
                    continue;
                }

                var element = new XElement(MainNs + "hyperlink", new XAttribute("ref", ToRangeReference(hyperlink)));
                if (!string.IsNullOrEmpty(hyperlink.Address))
                {
                    element.SetAttributeValue(RelationshipNs + "id", "rId" + relationshipId.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    relationshipId++;
                }

                if (!string.IsNullOrEmpty(hyperlink.SubAddress))
                {
                    element.SetAttributeValue("location", hyperlink.SubAddress);
                }

                if (!string.IsNullOrEmpty(hyperlink.ScreenTip))
                {
                    element.SetAttributeValue("tooltip", hyperlink.ScreenTip);
                }

                if (!string.IsNullOrEmpty(hyperlink.TextToDisplay))
                {
                    element.SetAttributeValue("display", hyperlink.TextToDisplay);
                }

                hyperlinksElement.Add(element);
            }

            return hyperlinksElement.IsEmpty ? null : hyperlinksElement;
        }

        internal static XDocument BuildWorksheetHyperlinkRelationships(WorksheetModel worksheet)
        {
            var orderedHyperlinks = GetOrderedHyperlinks(worksheet.Hyperlinks);
            var relationships = new XElement(PackageRelationshipNs + "Relationships");
            var relationshipId = 1;

            for (var index = 0; index < orderedHyperlinks.Count; index++)
            {
                var hyperlink = orderedHyperlinks[index];
                if (string.IsNullOrEmpty(hyperlink.Address))
                {
                    continue;
                }

                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + relationshipId.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                    new XAttribute("Type", HyperlinkRelationshipType),
                    new XAttribute("Target", hyperlink.Address),
                    new XAttribute("TargetMode", "External")));
                relationshipId++;
            }

            if (!relationships.HasElements)
            {
                return null;
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), relationships);
        }

        internal static Dictionary<string, string> LoadWorksheetHyperlinkTargets(ZipArchive archive, string worksheetUri)
        {
            var relationshipsEntry = GetEntry(archive, GetWorksheetRelationshipsUri(worksheetUri));
            if (relationshipsEntry == null)
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            var document = LoadDocument(relationshipsEntry);
            var targets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var relationship in document.Root?.Elements(PackageRelationshipNs + "Relationship") ?? Enumerable.Empty<XElement>())
            {
                var id = (string)relationship.Attribute("Id");
                var type = (string)relationship.Attribute("Type");
                var target = (string)relationship.Attribute("Target");
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target) || !string.Equals(type, HyperlinkRelationshipType, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                targets[id] = target;
            }

            return targets;
        }

        internal static void LoadHyperlinks(WorksheetModel worksheetModel, XElement worksheetRoot, IReadOnlyDictionary<string, string> hyperlinkTargets, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            worksheetModel.Hyperlinks.Clear();

            foreach (var hyperlinkElement in worksheetRoot.Element(MainNs + "hyperlinks")?.Elements(MainNs + "hyperlink") ?? Enumerable.Empty<XElement>())
            {
                var reference = (string)hyperlinkElement.Attribute("ref");
                HyperlinkModel candidate;
                if (!TryParseHyperlinkReference(reference, out candidate))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The hyperlink reference '" + reference + "' is invalid.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("HL-L001", DiagnosticSeverity.LossyRecoverable, "A hyperlink with an invalid reference was dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = reference,
                    });
                    continue;
                }

                if (ContainsOverlappingHyperlink(worksheetModel.Hyperlinks, candidate))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The hyperlink reference '" + reference + "' overlaps an existing hyperlink.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("HL-L003", DiagnosticSeverity.LossyRecoverable, "Overlapping hyperlink ranges were normalized during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = reference,
                    });
                    continue;
                }

                candidate.SubAddress = NormalizeAttributeValue((string)hyperlinkElement.Attribute("location"));
                candidate.ScreenTip = NormalizeAttributeValue((string)hyperlinkElement.Attribute("tooltip"));
                candidate.TextToDisplay = NormalizeAttributeValue((string)hyperlinkElement.Attribute("display"));

                var relationshipId = (string)hyperlinkElement.Attribute(RelationshipNs + "id");
                if (!string.IsNullOrEmpty(relationshipId))
                {
                    var resolvedRelationshipId = relationshipId;
                    string address;
                    if (hyperlinkTargets.TryGetValue(resolvedRelationshipId, out address))
                    {
                        candidate.Address = NormalizeAttributeValue(address);
                    }
                    else if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The hyperlink relationship '" + relationshipId + "' could not be resolved.");
                    }
                    else
                    {
                        AddIssue(diagnostics, options, new LoadIssue("HL-L002", DiagnosticSeverity.LossyRecoverable, "A hyperlink relationship was missing and the external address was cleared.", repairApplied: true, dataLossRisk: true)
                        {
                            SheetName = sheetName,
                            CellRef = reference,
                        });
                    }
                }

                if (string.IsNullOrEmpty(candidate.Address) && string.IsNullOrEmpty(candidate.SubAddress))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The hyperlink reference '" + reference + "' does not define an address or sub-address.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("HL-L004", DiagnosticSeverity.LossyRecoverable, "A hyperlink without an address or sub-address was dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = reference,
                    });
                    continue;
                }

                worksheetModel.Hyperlinks.Add(candidate);
            }
        }

        private static List<HyperlinkModel> GetOrderedHyperlinks(IReadOnlyList<HyperlinkModel> hyperlinks)
        {
            var ordered = new List<HyperlinkModel>(hyperlinks.Count);
            for (var index = 0; index < hyperlinks.Count; index++)
            {
                ordered.Add(hyperlinks[index]);
            }

            ordered.Sort(CompareHyperlinks);
            return ordered;
        }

        private static int CompareHyperlinks(HyperlinkModel left, HyperlinkModel right)
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

        private static bool ShouldPersistHyperlink(HyperlinkModel hyperlink)
        {
            return !string.IsNullOrEmpty(hyperlink.Address) || !string.IsNullOrEmpty(hyperlink.SubAddress);
        }

        private static string ToRangeReference(HyperlinkModel hyperlink)
        {
            var first = new CellAddress(hyperlink.FirstRow, hyperlink.FirstColumn).ToString();
            if (hyperlink.TotalRows == 1 && hyperlink.TotalColumns == 1)
            {
                return first;
            }

            var last = new CellAddress(hyperlink.FirstRow + hyperlink.TotalRows - 1, hyperlink.FirstColumn + hyperlink.TotalColumns - 1).ToString();
            return first + ":" + last;
        }

        private static string GetWorksheetRelationshipsUri(string worksheetUri)
        {
            var normalizedUri = worksheetUri.TrimStart('/');
            var slashIndex = normalizedUri.LastIndexOf('/');
            var directory = slashIndex >= 0 ? normalizedUri.Substring(0, slashIndex + 1) : string.Empty;
            var fileName = slashIndex >= 0 ? normalizedUri.Substring(slashIndex + 1) : normalizedUri;
            return "/" + directory + "_rels/" + fileName + ".rels";
        }

        private static bool TryParseHyperlinkReference(string reference, out HyperlinkModel hyperlink)
        {
            hyperlink = new HyperlinkModel();
            MergeRegion region;
            if (!TryParseMergeReference(reference ?? string.Empty, out region))
            {
                return false;
            }

            hyperlink.FirstRow = region.FirstRow;
            hyperlink.FirstColumn = region.FirstColumn;
            hyperlink.TotalRows = region.TotalRows;
            hyperlink.TotalColumns = region.TotalColumns;
            return true;
        }

        private static bool ContainsOverlappingHyperlink(IReadOnlyList<HyperlinkModel> hyperlinks, HyperlinkModel candidate)
        {
            for (var index = 0; index < hyperlinks.Count; index++)
            {
                if (HyperlinksOverlap(hyperlinks[index], candidate))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool HyperlinksOverlap(HyperlinkModel left, HyperlinkModel right)
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

        private static string NormalizeAttributeValue(string value)
        {
            return string.IsNullOrEmpty(value) ? null : value;
        }
    }
}
