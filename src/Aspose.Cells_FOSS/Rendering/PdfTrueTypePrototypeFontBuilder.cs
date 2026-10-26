using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfTrueTypePrototypeFontBuilder
    {
        public IDictionary<int, byte[]> BuildObjects(IList<PdfTrueTypePrototypeUsage> usages, ref int nextObjectId)
        {
            var objects = new Dictionary<int, byte[]>();
            for (var i = 0; i < usages.Count; i++)
            {
                var usage = usages[i];
                usage.ResourceName = "TTF" + (i + 1).ToString(CultureInfo.InvariantCulture);
                usage.SubsetFontName = BuildSubsetFontName(usage, i);
                usage.FontFileObjectId = nextObjectId++;
                usage.FontDescriptorObjectId = nextObjectId++;
                usage.ToUnicodeObjectId = nextObjectId++;
                usage.FontObjectId = nextObjectId++;

                var subsetBytes = usage.BuildSubsetFontBytes();
                objects[usage.FontFileObjectId] = BuildFontFileStreamObject(usage.FontFileObjectId, subsetBytes);
                objects[usage.FontDescriptorObjectId] = BuildFontDescriptorObject(usage);
                objects[usage.ToUnicodeObjectId] = BuildStreamObject(usage.ToUnicodeObjectId, BuildToUnicodeContent(usage));
                objects[usage.FontObjectId] = BuildFontObject(usage);
            }

            return objects;
        }

        private static byte[] BuildFontObject(PdfTrueTypePrototypeUsage usage)
        {
            var builder = new StringBuilder();
            builder.Append(usage.FontObjectId.ToString(CultureInfo.InvariantCulture));
            builder.Append(" 0 obj\n<< /Type /Font\n/Subtype /TrueType\n/Name /");
            builder.Append(usage.ResourceName);
            builder.Append("\n/BaseFont /");
            builder.Append(usage.SubsetFontName);
            builder.Append("\n/Encoding /WinAnsiEncoding\n/FontDescriptor ");
            builder.Append(usage.FontDescriptorObjectId.ToString(CultureInfo.InvariantCulture));
            builder.Append(" 0 R\n/FirstChar ");
            builder.Append(usage.FirstChar.ToString(CultureInfo.InvariantCulture));
            builder.Append("\n/LastChar ");
            builder.Append(usage.LastChar.ToString(CultureInfo.InvariantCulture));
            builder.Append("\n/Widths [");
            for (var code = usage.FirstChar; code <= usage.LastChar; code++)
            {
                if (code > usage.FirstChar)
                {
                    builder.Append(' ');
                }

                builder.Append(Format(usage.WidthForCode(code)));
            }

            builder.Append("]\n/ToUnicode ");
            builder.Append(usage.ToUnicodeObjectId.ToString(CultureInfo.InvariantCulture));
            builder.Append(" 0 R\n>>\nendobj\n");
            return Encoding.ASCII.GetBytes(builder.ToString());
        }

        private static byte[] BuildFontDescriptorObject(PdfTrueTypePrototypeUsage usage)
        {
            var builder = new StringBuilder();
            builder.Append(usage.FontDescriptorObjectId.ToString(CultureInfo.InvariantCulture));
            builder.Append(" 0 obj\n<< /Type /FontDescriptor\n/FontName /");
            builder.Append(usage.SubsetFontName);
            builder.Append("\n/Flags 32\n/ItalicAngle 0\n/Ascent 810\n/Descent -232\n/CapHeight 810\n/StemV 44\n/FontBBox [");
            if (usage.HasBounds)
            {
                builder.Append(Format(usage.FontBounds.Left));
                builder.Append(' ');
                builder.Append(Format(usage.FontBounds.Top));
                builder.Append(' ');
                builder.Append(Format(usage.FontBounds.Right));
                builder.Append(' ');
                builder.Append(Format(usage.FontBounds.Bottom));
            }
            else
            {
                builder.Append("0 0 0 0");
            }

            builder.Append("]\n/FontFile2 ");
            builder.Append(usage.FontFileObjectId.ToString(CultureInfo.InvariantCulture));
            builder.Append(" 0 R\n>>\nendobj\n");
            return Encoding.ASCII.GetBytes(builder.ToString());
        }

        private static byte[] BuildToUnicodeContent(PdfTrueTypePrototypeUsage usage)
        {
            var ordered = usage.OrderedCodes();
            var builder = new StringBuilder();
            builder.Append("/CIDInit /ProcSet findresource begin\n");
            builder.Append("12 dict begin\nbegincmap\n");
            builder.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
            builder.Append("/CMapName /");
            builder.Append(usage.ResourceName);
            builder.Append(" def\n/CMapType 2 def\n");
            builder.Append("1 begincodespacerange\n<00> <FF>\nendcodespacerange\n");
            builder.Append(ordered.Count.ToString(CultureInfo.InvariantCulture));
            builder.Append(" beginbfchar\n");
            for (var i = 0; i < ordered.Count; i++)
            {
                var code = ordered[i];
                builder.Append('<');
                builder.Append(code.ToString("X2", CultureInfo.InvariantCulture));
                builder.Append("> <00");
                builder.Append(code.ToString("X2", CultureInfo.InvariantCulture));
                builder.Append(">\n");
            }

            builder.Append("endbfchar\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
            return Encoding.ASCII.GetBytes(builder.ToString());
        }

        private static string BuildSubsetFontName(PdfTrueTypePrototypeUsage usage, int usageIndex)
        {
            var prefixValue = usageIndex + 1;
            var prefixChars = new char[6];
            for (var i = 5; i >= 0; i--)
            {
                prefixChars[i] = (char)('A' + (prefixValue % 26));
                prefixValue /= 26;
            }

            return new string(prefixChars) + "+" + usage.BaseFontName;
        }

        private static byte[] BuildStreamObject(int objectId, byte[] streamBytes)
        {
            var compressed = Compress(streamBytes);
            var header = objectId.ToString(CultureInfo.InvariantCulture)
                + " 0 obj\n<< /Filter /FlateDecode\n/Length "
                + compressed.Length.ToString(CultureInfo.InvariantCulture)
                + " >>\nstream\n";
            var footer = "\nendstream\nendobj\n";
            using (var output = new MemoryStream())
            {
                var headerBytes = Encoding.ASCII.GetBytes(header);
                var footerBytes = Encoding.ASCII.GetBytes(footer);
                output.Write(headerBytes, 0, headerBytes.Length);
                output.Write(compressed, 0, compressed.Length);
                output.Write(footerBytes, 0, footerBytes.Length);
                return output.ToArray();
            }
        }

        private static byte[] BuildFontFileStreamObject(int objectId, byte[] streamBytes)
        {
            var compressed = Compress(streamBytes);
            var header = objectId.ToString(CultureInfo.InvariantCulture)
                + " 0 obj\n<< /Filter /FlateDecode\n/Length "
                + compressed.Length.ToString(CultureInfo.InvariantCulture)
                + "\n/Length1 "
                + streamBytes.Length.ToString(CultureInfo.InvariantCulture)
                + " >>\nstream\n";
            var footer = "\nendstream\nendobj\n";
            using (var output = new MemoryStream())
            {
                var headerBytes = Encoding.ASCII.GetBytes(header);
                var footerBytes = Encoding.ASCII.GetBytes(footer);
                output.Write(headerBytes, 0, headerBytes.Length);
                output.Write(compressed, 0, compressed.Length);
                output.Write(footerBytes, 0, footerBytes.Length);
                return output.ToArray();
            }
        }

        private static byte[] Compress(byte[] raw)
        {
            using (var output = new MemoryStream())
            {
                using (var zip = new ZLibStream(output, CompressionLevel.Optimal, true))
                {
                    zip.Write(raw, 0, raw.Length);
                }

                return output.ToArray();
            }
        }

        private static string Format(float value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
