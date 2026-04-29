using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookComments
    {
        internal const string CommentsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        internal const string VmlDrawingRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
        internal const string CommentsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
        internal const string VmlDrawingContentType = "application/vnd.openxmlformats-officedocument.vmlDrawing";

        internal static XDocument BuildCommentsDocument(WorksheetModel worksheet)
        {
            var sorted = SortComments(worksheet.Comments);

            var authors = new List<string>();
            var authorToIndex = new Dictionary<string, int>(System.StringComparer.Ordinal);

            for (var i = 0; i < sorted.Count; i++)
            {
                var author = sorted[i].Author;
                if (author == null)
                {
                    author = string.Empty;
                }

                if (!authorToIndex.ContainsKey(author))
                {
                    authorToIndex[author] = authors.Count;
                    authors.Add(author);
                }
            }

            var authorsElement = new XElement(MainNs + "authors");
            for (var a = 0; a < authors.Count; a++)
            {
                authorsElement.Add(new XElement(MainNs + "author", authors[a]));
            }

            var commentListElement = new XElement(MainNs + "commentList");
            for (var i = 0; i < sorted.Count; i++)
            {
                var comment = sorted[i];
                var author = comment.Author == null ? string.Empty : comment.Author;
                int authorId;
                authorToIndex.TryGetValue(author, out authorId);

                var cellRef = new CellAddress(comment.Row, comment.Column).ToString();
                var note = comment.Note == null ? string.Empty : comment.Note;

                var tElement = new XElement(MainNs + "t", note);
                if (NeedsPreserveWhitespace(note))
                {
                    tElement.SetAttributeValue(XmlNs + "space", "preserve");
                }

                var commentElement = new XElement(MainNs + "comment",
                    new XAttribute("ref", cellRef),
                    new XAttribute("authorId", authorId.ToString(CultureInfo.InvariantCulture)),
                    new XElement(MainNs + "text",
                        new XElement(MainNs + "r", tElement)));

                commentListElement.Add(commentElement);
            }

            var root = new XElement(MainNs + "comments", authorsElement, commentListElement);
            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
        }

        internal static void WriteVmlDrawing(ZipArchive archive, WorksheetModel worksheet, int commentFileNumber)
        {
            var path = "xl/drawings/vmlDrawing" + commentFileNumber.ToString(CultureInfo.InvariantCulture) + ".vml";
            var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(BuildVmlContent(worksheet));
            }
        }

        private static string BuildVmlContent(WorksheetModel worksheet)
        {
            var sb = new StringBuilder();
            sb.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"");
            sb.Append(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            sb.Append(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
            sb.Append("<o:shapelayout v:ext=\"edit\">");
            sb.Append("<o:idmap v:ext=\"edit\" data=\"1\"/>");
            sb.Append("</o:shapelayout>");
            sb.Append("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"");
            sb.Append(" path=\"m,l,21600r21600,l21600,xe\">");
            sb.Append("<v:stroke joinstyle=\"miter\"/>");
            sb.Append("<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>");
            sb.Append("</v:shapetype>");

            var sorted = SortComments(worksheet.Comments);
            for (var i = 0; i < sorted.Count; i++)
            {
                var comment = sorted[i];
                var shapeId = 1025 + i;
                if (comment.RawVmlShapeXml != null)
                {
                    sb.Append(PatchVmlShape(comment.RawVmlShapeXml, comment, shapeId));
                }
                else
                {
                    sb.Append(BuildDefaultVmlShape(comment, shapeId));
                }
            }

            sb.Append("</xml>");
            return sb.ToString();
        }

        private static string PatchVmlShape(string raw, CommentModel comment, int shapeId)
        {
            raw = ReplaceInlineValue(raw, "id=\"_x0000_s", "\"", shapeId.ToString(CultureInfo.InvariantCulture));

            var widthPt = FormatPt(comment.Width);
            var heightPt = FormatPt(comment.Height);
            raw = ReplaceInlineValue(raw, "width:", "pt", widthPt);
            raw = ReplaceInlineValue(raw, "height:", "pt", heightPt);

            raw = SetStyleVisibility(raw, comment.IsVisible);
            raw = ReplaceTagContent(raw, "x:Row", comment.Row.ToString(CultureInfo.InvariantCulture));
            raw = ReplaceTagContent(raw, "x:Column", comment.Column.ToString(CultureInfo.InvariantCulture));
            raw = SetClientDataVisible(raw, comment.IsVisible);

            return raw;
        }

        private static string BuildDefaultVmlShape(CommentModel comment, int shapeId)
        {
            var widthPt = FormatPt(comment.Width);
            var heightPt = FormatPt(comment.Height);
            var visStyle = comment.IsVisible ? ";visibility:visible" : string.Empty;

            var anchorLeft = comment.Column + 1;
            var anchorTop = comment.Row > 0 ? comment.Row - 1 : 0;
            var anchorRight = comment.Column + 4;
            var anchorBottom = comment.Row + 5;

            var sb = new StringBuilder();
            sb.Append("<v:shape id=\"_x0000_s");
            sb.Append(shapeId.ToString(CultureInfo.InvariantCulture));
            sb.Append("\" type=\"#_x0000_t202\"");
            sb.Append(" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:");
            sb.Append(widthPt);
            sb.Append("pt;height:");
            sb.Append(heightPt);
            sb.Append("pt;z-index:1");
            sb.Append(visStyle);
            sb.Append("\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">");
            sb.Append("<v:fill color2=\"#ffffe1\"/>");
            sb.Append("<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
            sb.Append("<v:path o:connecttype=\"none\"/>");
            sb.Append("<v:textbox style=\"mso-direction-alt:auto\">");
            sb.Append("<div style=\"text-align:left\"/>");
            sb.Append("</v:textbox>");
            sb.Append("<x:ClientData ObjectType=\"Note\">");
            sb.Append("<x:MoveWithCells/>");
            sb.Append("<x:SizeWithCells/>");
            sb.Append("<x:Anchor>");
            sb.Append(anchorLeft.ToString(CultureInfo.InvariantCulture));
            sb.Append(",15,");
            sb.Append(anchorTop.ToString(CultureInfo.InvariantCulture));
            sb.Append(",2,");
            sb.Append(anchorRight.ToString(CultureInfo.InvariantCulture));
            sb.Append(",15,");
            sb.Append(anchorBottom.ToString(CultureInfo.InvariantCulture));
            sb.Append(",16</x:Anchor>");
            sb.Append("<x:AutoFill>False</x:AutoFill>");
            sb.Append("<x:Row>");
            sb.Append(comment.Row.ToString(CultureInfo.InvariantCulture));
            sb.Append("</x:Row>");
            sb.Append("<x:Column>");
            sb.Append(comment.Column.ToString(CultureInfo.InvariantCulture));
            sb.Append("</x:Column>");
            if (comment.IsVisible)
            {
                sb.Append("<x:Visible/>");
            }

            sb.Append("</x:ClientData>");
            sb.Append("</v:shape>");
            return sb.ToString();
        }

        internal static string FormatPt(int pixels)
        {
            var pt = pixels * 0.75;
            return pt.ToString("0.##", CultureInfo.InvariantCulture);
        }

        internal static List<CommentModel> SortComments(List<CommentModel> comments)
        {
            var sorted = new List<CommentModel>(comments);
            for (var i = 0; i < sorted.Count - 1; i++)
            {
                for (var j = i + 1; j < sorted.Count; j++)
                {
                    var a = sorted[i];
                    var b = sorted[j];
                    if (a.Row > b.Row || (a.Row == b.Row && a.Column > b.Column))
                    {
                        sorted[i] = b;
                        sorted[j] = a;
                    }
                }
            }

            return sorted;
        }

        // Replaces the value between prefix and the next occurrence of suffix.
        internal static string ReplaceInlineValue(string text, string prefix, string suffix, string newValue)
        {
            var start = text.IndexOf(prefix, System.StringComparison.Ordinal);
            if (start < 0)
            {
                return text;
            }

            var valueStart = start + prefix.Length;
            var end = text.IndexOf(suffix, valueStart, System.StringComparison.Ordinal);
            if (end < 0)
            {
                return text;
            }

            return text.Substring(0, valueStart) + newValue + text.Substring(end);
        }

        // Replaces the text content of <tagName>...</tagName> (first occurrence).
        internal static string ReplaceTagContent(string text, string tagName, string newContent)
        {
            var openTag = "<" + tagName + ">";
            var closeTag = "</" + tagName + ">";
            var start = text.IndexOf(openTag, System.StringComparison.Ordinal);
            if (start < 0)
            {
                return text;
            }

            var contentStart = start + openTag.Length;
            var end = text.IndexOf(closeTag, contentStart, System.StringComparison.Ordinal);
            if (end < 0)
            {
                return text;
            }

            return text.Substring(0, contentStart) + newContent + text.Substring(end);
        }

        private static string SetStyleVisibility(string raw, bool isVisible)
        {
            bool hasVisible = raw.IndexOf("visibility:visible", System.StringComparison.Ordinal) >= 0;

            if (isVisible && !hasVisible)
            {
                var stylePrefix = "style=\"";
                var styleStart = raw.IndexOf(stylePrefix, System.StringComparison.Ordinal);
                if (styleStart >= 0)
                {
                    var quoteStart = styleStart + stylePrefix.Length;
                    var quoteEnd = raw.IndexOf('"', quoteStart);
                    if (quoteEnd >= 0)
                    {
                        return raw.Substring(0, quoteEnd) + ";visibility:visible" + raw.Substring(quoteEnd);
                    }
                }

                return raw;
            }

            if (!isVisible && hasVisible)
            {
                var withSemi = ";visibility:visible";
                var idx = raw.IndexOf(withSemi, System.StringComparison.Ordinal);
                if (idx >= 0)
                {
                    return raw.Remove(idx, withSemi.Length);
                }

                var noSemi = "visibility:visible";
                idx = raw.IndexOf(noSemi, System.StringComparison.Ordinal);
                if (idx >= 0)
                {
                    return raw.Remove(idx, noSemi.Length);
                }
            }

            return raw;
        }

        private static string SetClientDataVisible(string raw, bool isVisible)
        {
            bool hasTag = raw.IndexOf("<x:Visible/>", System.StringComparison.Ordinal) >= 0;

            if (isVisible && !hasTag)
            {
                var endTag = "</x:ClientData>";
                var idx = raw.IndexOf(endTag, System.StringComparison.Ordinal);
                if (idx >= 0)
                {
                    return raw.Substring(0, idx) + "<x:Visible/>" + raw.Substring(idx);
                }
            }
            else if (!isVisible && hasTag)
            {
                return raw.Replace("<x:Visible/>", string.Empty);
            }

            return raw;
        }
    }
}
