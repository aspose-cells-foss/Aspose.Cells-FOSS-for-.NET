using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfObjectTable
    {
        private readonly Dictionary<int, byte[]> _objects = new Dictionary<int, byte[]>();
        private readonly List<int> _pageObjectIds = new List<int>();

        public byte[] HeaderBytes;
        public string RootReference;
        public string InfoReference;

        public IDictionary<int, byte[]> Objects
        {
            get { return _objects; }
        }

        public IList<int> PageObjectIds
        {
            get { return _pageObjectIds; }
        }

        public int MaxObjectId
        {
            get
            {
                var max = 0;
                foreach (var key in _objects.Keys)
                {
                    if (key > max)
                    {
                        max = key;
                    }
                }

                return max;
            }
        }

        public static PdfObjectTable Parse(byte[] pdf)
        {
            var text = Latin1(pdf);
            var startXrefIndex = text.LastIndexOf("startxref", StringComparison.Ordinal);
            if (startXrefIndex < 0)
            {
                throw new InvalidOperationException("PDF startxref not found.");
            }

            var xrefOffset = ParseStartXref(text, startXrefIndex);
            var trailerIndex = text.IndexOf("trailer", xrefOffset, StringComparison.Ordinal);
            if (trailerIndex < 0)
            {
                throw new InvalidOperationException("PDF trailer not found.");
            }

            var table = new PdfObjectTable();
            var offsets = ParseXrefOffsets(text, xrefOffset, trailerIndex);
            var sorted = new List<int>();
            foreach (var pair in offsets)
            {
                if (pair.Key > 0)
                {
                    sorted.Add(pair.Key);
                }
            }

            sorted.Sort(delegate (int a, int b) { return offsets[a].CompareTo(offsets[b]); });
            var firstOffset = int.MaxValue;
            for (var i = 0; i < sorted.Count; i++)
            {
                if (offsets[sorted[i]] < firstOffset)
                {
                    firstOffset = offsets[sorted[i]];
                }
            }

            table.HeaderBytes = new byte[firstOffset];
            Buffer.BlockCopy(pdf, 0, table.HeaderBytes, 0, firstOffset);

            for (var i = 0; i < sorted.Count; i++)
            {
                var objectId = sorted[i];
                var start = offsets[objectId];
                var end = i + 1 < sorted.Count ? offsets[sorted[i + 1]] : xrefOffset;
                var length = end - start;
                if (length <= 0)
                {
                    continue;
                }

                var bytes = new byte[length];
                Buffer.BlockCopy(pdf, start, bytes, 0, length);
                table._objects[objectId] = bytes;

                var objectText = Latin1(bytes);
                if (objectText.IndexOf("/Type /Page", StringComparison.Ordinal) >= 0
                    && objectText.IndexOf("/Type /Pages", StringComparison.Ordinal) < 0)
                {
                    table._pageObjectIds.Add(objectId);
                }
            }

            var trailer = text.Substring(trailerIndex, startXrefIndex - trailerIndex);
            table.RootReference = ParseReference(trailer, "/Root");
            table.InfoReference = ParseReference(trailer, "/Info");
            return table;
        }

        private static Dictionary<int, int> ParseXrefOffsets(string text, int xrefOffset, int trailerIndex)
        {
            var offsets = new Dictionary<int, int>();
            using (var reader = new StringReader(text.Substring(xrefOffset, trailerIndex - xrefOffset)))
            {
                var first = reader.ReadLine();
                if (first == null || first.Trim() != "xref")
                {
                    throw new InvalidOperationException("Unsupported PDF xref table.");
                }

                while (true)
                {
                    var line = reader.ReadLine();
                    if (line == null)
                    {
                        break;
                    }

                    line = line.Trim();
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length != 2)
                    {
                        break;
                    }

                    var start = int.Parse(parts[0], CultureInfo.InvariantCulture);
                    var count = int.Parse(parts[1], CultureInfo.InvariantCulture);
                    for (var i = 0; i < count; i++)
                    {
                        var entry = reader.ReadLine();
                        if (entry == null)
                        {
                            throw new InvalidOperationException("Unexpected end of xref table.");
                        }

                        if (entry.Length < 17)
                        {
                            continue;
                        }

                        if (entry[17] == 'n')
                        {
                            offsets[start + i] = int.Parse(entry.Substring(0, 10), CultureInfo.InvariantCulture);
                        }
                    }
                }
            }

            return offsets;
        }

        private static int ParseStartXref(string text, int startXrefIndex)
        {
            var cursor = startXrefIndex + "startxref".Length;
            while (cursor < text.Length && char.IsWhiteSpace(text[cursor]))
            {
                cursor++;
            }

            var end = cursor;
            while (end < text.Length && char.IsDigit(text[end]))
            {
                end++;
            }

            return int.Parse(text.Substring(cursor, end - cursor), CultureInfo.InvariantCulture);
        }

        private static string ParseReference(string trailer, string key)
        {
            var index = trailer.IndexOf(key, StringComparison.Ordinal);
            if (index < 0)
            {
                return null;
            }

            index += key.Length;
            while (index < trailer.Length && char.IsWhiteSpace(trailer[index]))
            {
                index++;
            }

            var end = index;
            while (end < trailer.Length && trailer[end] != '\n' && trailer[end] != '\r' && trailer[end] != '>')
            {
                end++;
            }

            return trailer.Substring(index, end - index).Trim();
        }

        private static string Latin1(byte[] bytes)
        {
            var chars = new char[bytes.Length];
            for (var i = 0; i < bytes.Length; i++)
            {
                chars[i] = (char)bytes[i];
            }

            return new string(chars);
        }
    }
}
