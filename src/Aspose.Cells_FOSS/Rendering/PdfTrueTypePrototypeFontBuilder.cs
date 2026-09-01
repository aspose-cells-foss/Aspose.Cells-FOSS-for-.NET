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
            var zlibStreamType = System.Type.GetType("System.IO.Compression.ZLibStream, System.IO.Compression", false);
            if (zlibStreamType != null)
            {
                var constructor = zlibStreamType.GetConstructor(new[] { typeof(Stream), typeof(CompressionLevel), typeof(bool) });
                if (constructor != null)
                {
                    using (var output = new MemoryStream())
                    {
                        System.IDisposable stream = null;
                        try
                        {
                            stream = (System.IDisposable)constructor.Invoke(new object[] { output, CompressionLevel.Optimal, true });
                            ((Stream)stream).Write(raw, 0, raw.Length);
                        }
                        finally
                        {
                            if (stream != null)
                            {
                                stream.Dispose();
                            }
                        }

                        return output.ToArray();
                    }
                }
            }

            using (var output = new MemoryStream())
            {
                output.WriteByte(0x78);
                output.WriteByte(0xDA);
                using (var deflated = new MemoryStream())
                {
                    using (var zip = new DeflateStream(deflated, CompressionLevel.Optimal, true))
                    {
                        zip.Write(raw, 0, raw.Length);
                    }

                    var body = deflated.ToArray();
                    output.Write(body, 0, body.Length);
                }

                var adler = Adler32(raw);
                output.WriteByte((byte)((adler >> 24) & 0xFF));
                output.WriteByte((byte)((adler >> 16) & 0xFF));
                output.WriteByte((byte)((adler >> 8) & 0xFF));
                output.WriteByte((byte)(adler & 0xFF));
                return output.ToArray();
            }
        }

        private static uint Adler32(byte[] data)
        {
            const uint mod = 65521;
            uint a = 1;
            uint b = 0;

            for (var i = 0; i < data.Length; i++)
            {
                a = (a + data[i]) % mod;
                b = (b + a) % mod;
            }

            return (b << 16) | a;
        }

        private static string Format(float value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
