using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookComments;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookCommentsLoader
    {
        internal static void LoadComments(WorksheetModel worksheetModel, ZipArchive archive, string worksheetUri, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            worksheetModel.Comments.Clear();

            var relsUri = GetWorksheetRelsUri(worksheetUri);
            var relsEntry = GetEntry(archive, relsUri);
            if (relsEntry == null)
            {
                return;
            }

            System.Xml.Linq.XDocument relsDoc;
            try
            {
                relsDoc = LoadDocument(relsEntry);
            }
            catch
            {
                return;
            }

            if (relsDoc.Root == null)
            {
                return;
            }

            string commentsTarget = null;
            string vmlTarget = null;

            foreach (var rel in relsDoc.Root.Elements(PackageRelationshipNs + "Relationship"))
            {
                var type = (string)rel.Attribute("Type");
                var target = (string)rel.Attribute("Target");
                if (string.IsNullOrEmpty(type) || string.IsNullOrEmpty(target))
                {
                    continue;
                }

                if (string.Equals(type, CommentsRelationshipType, StringComparison.OrdinalIgnoreCase))
                {
                    commentsTarget = ResolvePartUri(worksheetUri, target);
                }
                else if (string.Equals(type, VmlDrawingRelationshipType, StringComparison.OrdinalIgnoreCase))
                {
                    vmlTarget = ResolvePartUri(worksheetUri, target);
                }
            }

            if (commentsTarget == null)
            {
                if (vmlTarget != null)
                {
                    AddIssue(diagnostics, options, new LoadIssue("COM-R003", DiagnosticSeverity.Recoverable, "A VML drawing part was found without a comments relationship; comments were skipped.", repairApplied: false)
                    {
                        SheetName = sheetName,
                    });
                }

                return;
            }

            var commentsEntry = GetEntry(archive, commentsTarget);
            if (commentsEntry == null)
            {
                AddIssue(diagnostics, options, new LoadIssue("COM-R002", DiagnosticSeverity.Recoverable, "Comments part '" + commentsTarget + "' was referenced but not found.", repairApplied: false)
                {
                    SheetName = sheetName,
                });
                return;
            }

            LoadCommentsXml(worksheetModel, commentsEntry, diagnostics, options, sheetName);

            if (vmlTarget == null)
            {
                AddIssue(diagnostics, options, new LoadIssue("COM-R002", DiagnosticSeverity.Recoverable, "Comments were loaded but the VML drawing part is missing; default shape dimensions will be used.", repairApplied: true)
                {
                    SheetName = sheetName,
                });
                return;
            }

            var vmlEntry = GetEntry(archive, vmlTarget);
            if (vmlEntry == null)
            {
                AddIssue(diagnostics, options, new LoadIssue("COM-R002", DiagnosticSeverity.Recoverable, "VML drawing part '" + vmlTarget + "' was referenced but not found; default shape dimensions will be used.", repairApplied: true)
                {
                    SheetName = sheetName,
                });
                return;
            }

            LoadVmlDrawing(worksheetModel, vmlEntry);
        }

        private static void LoadCommentsXml(WorksheetModel worksheetModel, ZipArchiveEntry entry, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            System.Xml.Linq.XDocument document;
            try
            {
                document = LoadDocument(entry);
            }
            catch
            {
                return;
            }

            if (document.Root == null)
            {
                return;
            }

            var authors = new List<string>();
            var authorsElement = document.Root.Element(MainNs + "authors");
            if (authorsElement != null)
            {
                foreach (var authorElement in authorsElement.Elements(MainNs + "author"))
                {
                    authors.Add(authorElement.Value ?? string.Empty);
                }
            }

            var commentListElement = document.Root.Element(MainNs + "commentList");
            if (commentListElement == null)
            {
                return;
            }

            foreach (var commentElement in commentListElement.Elements(MainNs + "comment"))
            {
                var cellRef = (string)commentElement.Attribute("ref");
                if (string.IsNullOrWhiteSpace(cellRef))
                {
                    AddIssue(diagnostics, options, new LoadIssue("COM-R001", DiagnosticSeverity.Recoverable, "A comment with a missing ref attribute was skipped.", repairApplied: false)
                    {
                        SheetName = sheetName,
                    });
                    continue;
                }

                CellAddress address;
                if (!TryParseCellReference(cellRef, out address))
                {
                    AddIssue(diagnostics, options, new LoadIssue("COM-R001", DiagnosticSeverity.Recoverable, "A comment with invalid ref '" + cellRef + "' was skipped.", repairApplied: false)
                    {
                        SheetName = sheetName,
                        CellRef = cellRef,
                    });
                    continue;
                }

                var author = string.Empty;
                var authorIdStr = (string)commentElement.Attribute("authorId");
                int authorId;
                if (!string.IsNullOrEmpty(authorIdStr)
                    && int.TryParse(authorIdStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out authorId)
                    && authorId >= 0
                    && authorId < authors.Count)
                {
                    author = authors[authorId];
                }

                var textElement = commentElement.Element(MainNs + "text");
                var note = ReadNoteText(textElement);

                var model = new CommentModel();
                model.Row = address.RowIndex;
                model.Column = address.ColumnIndex;
                model.Author = author;
                model.Note = note;
                worksheetModel.Comments.Add(model);
            }
        }

        private static string ReadNoteText(System.Xml.Linq.XElement textElement)
        {
            if (textElement == null)
            {
                return string.Empty;
            }

            var hasRuns = false;
            var sb = new StringBuilder();
            foreach (var runElement in textElement.Elements(MainNs + "r"))
            {
                hasRuns = true;
                var tElement = runElement.Element(MainNs + "t");
                if (tElement != null)
                {
                    sb.Append(tElement.Value);
                }
            }

            if (!hasRuns)
            {
                return textElement.Value;
            }

            return sb.ToString();
        }

        private static void LoadVmlDrawing(WorksheetModel worksheetModel, ZipArchiveEntry entry)
        {
            string vmlText;
            try
            {
                using (var rawStream = entry.Open())
                using (var reader = new StreamReader(rawStream, Encoding.UTF8))
                {
                    vmlText = reader.ReadToEnd();
                }
            }
            catch
            {
                return;
            }

            var commentMap = new Dictionary<long, CommentModel>();
            for (var i = 0; i < worksheetModel.Comments.Count; i++)
            {
                var c = worksheetModel.Comments[i];
                var key = ((long)c.Row << 20) | (long)c.Column;
                if (!commentMap.ContainsKey(key))
                {
                    commentMap[key] = c;
                }
            }

            var shapes = ExtractVmlShapes(vmlText);
            for (var s = 0; s < shapes.Count; s++)
            {
                var shapeText = shapes[s];

                if (shapeText.IndexOf("ObjectType=\"Note\"", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                var row = ExtractTagInt(shapeText, "x:Row");
                var col = ExtractTagInt(shapeText, "x:Column");
                if (row < 0 || col < 0)
                {
                    continue;
                }

                var mapKey = ((long)row << 20) | (long)col;
                CommentModel model;
                if (!commentMap.TryGetValue(mapKey, out model))
                {
                    continue;
                }

                var styleAttr = ExtractAttributeValue(shapeText, "style=\"");
                if (styleAttr != null)
                {
                    model.IsVisible = styleAttr.IndexOf("visibility:visible", StringComparison.OrdinalIgnoreCase) >= 0;

                    var w = ParseStylePt(styleAttr, "width:");
                    if (w > 0)
                    {
                        model.Width = w;
                    }

                    var h = ParseStylePt(styleAttr, "height:");
                    if (h > 0)
                    {
                        model.Height = h;
                    }
                }

                model.RawVmlShapeXml = shapeText;
            }
        }

        private static List<string> ExtractVmlShapes(string vmlText)
        {
            var shapes = new List<string>();
            var searchStart = 0;
            var closeTag = "</v:shape>";

            while (true)
            {
                var shapeStart = FindNextVmlShape(vmlText, searchStart);
                if (shapeStart < 0)
                {
                    break;
                }

                var shapeEnd = vmlText.IndexOf(closeTag, shapeStart, StringComparison.Ordinal);
                if (shapeEnd < 0)
                {
                    break;
                }

                var endPos = shapeEnd + closeTag.Length;
                shapes.Add(vmlText.Substring(shapeStart, endPos - shapeStart));
                searchStart = endPos;
            }

            return shapes;
        }

        // Finds the next <v:shape> start (not <v:shapetype> or any other v:shapeXxx variant).
        private static int FindNextVmlShape(string vmlText, int searchFrom)
        {
            var pos = searchFrom;
            while (true)
            {
                var idx = vmlText.IndexOf("<v:shape", pos, StringComparison.Ordinal);
                if (idx < 0)
                {
                    return -1;
                }

                var afterTag = idx + 8; // length of "<v:shape"
                if (afterTag >= vmlText.Length)
                {
                    return -1;
                }

                var next = vmlText[afterTag];
                if (next == ' ' || next == '>' || next == '\r' || next == '\n' || next == '\t')
                {
                    return idx;
                }

                pos = idx + 1;
            }
        }

        private static int ExtractTagInt(string text, string tagName)
        {
            var openTag = "<" + tagName + ">";
            var closeTag = "</" + tagName + ">";
            var start = text.IndexOf(openTag, StringComparison.Ordinal);
            if (start < 0)
            {
                return -1;
            }

            var valueStart = start + openTag.Length;
            var end = text.IndexOf(closeTag, valueStart, StringComparison.Ordinal);
            if (end < 0)
            {
                return -1;
            }

            int result;
            var valuePart = text.Substring(valueStart, end - valueStart).Trim();
            if (int.TryParse(valuePart, NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }

            return -1;
        }

        private static string ExtractAttributeValue(string text, string prefix)
        {
            var start = text.IndexOf(prefix, StringComparison.Ordinal);
            if (start < 0)
            {
                return null;
            }

            var valueStart = start + prefix.Length;
            var end = text.IndexOf('"', valueStart);
            if (end < 0)
            {
                return null;
            }

            return text.Substring(valueStart, end - valueStart);
        }

        private static int ParseStylePt(string style, string prefix)
        {
            var start = style.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
            if (start < 0)
            {
                return 0;
            }

            var valueStart = start + prefix.Length;
            var end = style.IndexOf(';', valueStart);
            if (end < 0)
            {
                end = style.IndexOf('"', valueStart);
            }

            if (end < 0)
            {
                end = style.Length;
            }

            var valuePart = style.Substring(valueStart, end - valueStart).Trim();
            if (valuePart.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                valuePart = valuePart.Substring(0, valuePart.Length - 2);
            }

            double pt;
            if (double.TryParse(valuePart, NumberStyles.Float, CultureInfo.InvariantCulture, out pt) && pt > 0)
            {
                return (int)Math.Round(pt / 0.75);
            }

            return 0;
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
