using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfType3FontBuilder
    {
        public Dictionary<int, byte[]> BuildObjects(PdfFontUsageRegistry registry, ref int nextObjectId)
        {
            var objects = new Dictionary<int, byte[]>();

            for (var i = 0; i < registry.Usages.Count; i++)
            {
                var usage = registry.Usages[i];
                for (var s = 0; s < usage.Subsets.Count; s++)
                {
                    var subset = usage.Subsets[s];
                    var glyphObjectIds = new Dictionary<string, int>(System.StringComparer.Ordinal);
                    var fontBounds = SKRect.Empty;
                    var fontBoundsInitialized = false;

                    for (var t = 0; t < subset.Tokens.Count; t++)
                    {
                        var token = subset.Tokens[t];
                        var width = usage.GetTokenWidth(token);
                        using (var path = usage.BuildTokenPath(token))
                        {
                            var content = BuildCharProcContent(path, width);
                            var objectId = nextObjectId++;
                            glyphObjectIds[token] = objectId;
                            objects[objectId] = BuildStreamObject(objectId, content);

                            if (path != null && !path.IsEmpty)
                            {
                                var bounds = path.Bounds;
                                if (!fontBoundsInitialized)
                                {
                                    fontBounds = bounds;
                                    fontBoundsInitialized = true;
                                }
                                else
                                {
                                    fontBounds = SKRect.Union(fontBounds, bounds);
                                }
                            }
                        }
                    }

                    subset.ToUnicodeObjectId = nextObjectId++;
                    objects[subset.ToUnicodeObjectId] = BuildStreamObject(subset.ToUnicodeObjectId, BuildToUnicodeContent(subset));

                    subset.UsePositiveYFontMatrix = usage.UsePositiveYFontMatrix;
                    subset.FontObjectId = nextObjectId++;
                    objects[subset.FontObjectId] = BuildFontObject(subset, usage, glyphObjectIds, fontBounds, fontBoundsInitialized);
                }
            }

            return objects;
        }

        private static byte[] BuildFontObject(PdfType3FontSubset subset, PdfType3FontUsage usage, Dictionary<string, int> glyphObjectIds, SKRect fontBounds, bool hasBounds)
        {
            var builder = new StringBuilder();
            builder.Append(subset.FontObjectId.ToString(CultureInfo.InvariantCulture));
            builder.Append(" 0 obj\n<< /Type /Font\n/Subtype /Type3\n/Name /");
            builder.Append(subset.ResourceName);
            builder.Append("\n/FontBBox [");
            if (hasBounds)
            {
                builder.Append(Format(fontBounds.Left));
                builder.Append(' ');
                builder.Append(Format(fontBounds.Top));
                builder.Append(' ');
                builder.Append(Format(fontBounds.Right));
                builder.Append(' ');
                builder.Append(Format(fontBounds.Bottom));
            }
            else
            {
                builder.Append("0 0 0 0");
            }

            builder.Append("]\n/FontMatrix ");
            builder.Append(BuildFontMatrix(subset.UsePositiveYFontMatrix));
            builder.Append("\n/CharProcs <<");

            for (var i = 0; i < subset.Tokens.Count; i++)
            {
                var token = subset.Tokens[i];
                var code = subset.TokenCodes[token];
                builder.Append(" /g");
                builder.Append(code.ToString(CultureInfo.InvariantCulture));
                builder.Append(' ');
                builder.Append(glyphObjectIds[token].ToString(CultureInfo.InvariantCulture));
                builder.Append(" 0 R");
            }

            builder.Append(" >>\n/Encoding << /Type /Encoding /Differences [1");
            for (var i = 0; i < subset.Tokens.Count; i++)
            {
                var token = subset.Tokens[i];
                var code = subset.TokenCodes[token];
                builder.Append(" /g");
                builder.Append(code.ToString(CultureInfo.InvariantCulture));
            }

            builder.Append("] >>\n/FirstChar 1\n/LastChar ");
            builder.Append(subset.Tokens.Count.ToString(CultureInfo.InvariantCulture));
            builder.Append("\n/Widths [");
            for (var i = 0; i < subset.Tokens.Count; i++)
            {
                if (i > 0)
                {
                    builder.Append(' ');
                }

                builder.Append(Format(usage.GetTokenWidth(subset.Tokens[i])));
            }

            builder.Append("]\n/Resources << >>\n/ToUnicode ");
            builder.Append(subset.ToUnicodeObjectId.ToString(CultureInfo.InvariantCulture));
            builder.Append(" 0 R\n>>\nendobj\n");
            return Encoding.ASCII.GetBytes(builder.ToString());
        }

        private static byte[] BuildCharProcContent(SKPath path, float width)
        {
            var builder = new StringBuilder();
            if (path == null || path.IsEmpty)
            {
                builder.Append(Format(width));
                builder.Append(" 0 d0\n");
                return Encoding.ASCII.GetBytes(builder.ToString());
            }

            var bounds = path.Bounds;
            builder.Append(Format(width));
            builder.Append(" 0 ");
            builder.Append(Format(bounds.Left));
            builder.Append(' ');
            builder.Append(Format(bounds.Top));
            builder.Append(' ');
            builder.Append(Format(bounds.Right));
            builder.Append(' ');
            builder.Append(Format(bounds.Bottom));
            builder.Append(" d1\n");
            builder.Append(PdfPathSerializer.Serialize(path));
            builder.Append("f\n");
            return Encoding.ASCII.GetBytes(builder.ToString());
        }

        private static byte[] BuildToUnicodeContent(PdfType3FontSubset subset)
        {
            var builder = new StringBuilder();
            var ranges = new List<PdfToUnicodeRange>();
            var singles = new List<PdfToUnicodeSingle>();
            BuildToUnicodeMappings(subset, ranges, singles);

            builder.Append("/CIDInit /ProcSet findresource begin\n");
            builder.Append("12 dict begin\nbegincmap\n");
            builder.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
            builder.Append("/CMapName /");
            builder.Append(subset.ResourceName);
            builder.Append(" def\n/CMapType 2 def\n");
            builder.Append("1 begincodespacerange\n<00> <FF>\nendcodespacerange\n");

            if (ranges.Count > 0)
            {
                builder.Append(ranges.Count.ToString(CultureInfo.InvariantCulture));
                builder.Append(" beginbfrange\n");
                for (var i = 0; i < ranges.Count; i++)
                {
                    builder.Append('<');
                    builder.Append(ranges[i].StartCode.ToString("X2", CultureInfo.InvariantCulture));
                    builder.Append("> <");
                    builder.Append(ranges[i].EndCode.ToString("X2", CultureInfo.InvariantCulture));
                    builder.Append("> <");
                    builder.Append(ranges[i].StartUnicode.ToString("X4", CultureInfo.InvariantCulture));
                    builder.Append(">\n");
                }

                builder.Append("endbfrange\n");
            }

            if (singles.Count > 0)
            {
                builder.Append(singles.Count.ToString(CultureInfo.InvariantCulture));
                builder.Append(" beginbfchar\n");
                for (var i = 0; i < singles.Count; i++)
                {
                    builder.Append('<');
                    builder.Append(singles[i].Code.ToString("X2", CultureInfo.InvariantCulture));
                    builder.Append("> <");
                    builder.Append(ToHexUnicode(singles[i].Token));
                    builder.Append(">\n");
                }

                builder.Append("endbfchar\n");
            }

            builder.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
            return Encoding.ASCII.GetBytes(builder.ToString());
        }

        private static void BuildToUnicodeMappings(PdfType3FontSubset subset, List<PdfToUnicodeRange> ranges, List<PdfToUnicodeSingle> singles)
        {
            var index = 0;
            while (index < subset.Tokens.Count)
            {
                var token = subset.Tokens[index];
                var code = subset.TokenCodes[token];
                var unicode = SingleUnicodeScalar(token);
                if (unicode < 0)
                {
                    singles.Add(new PdfToUnicodeSingle(code, token));
                    index++;
                    continue;
                }

                var endIndex = index;
                var endCode = code;
                var endUnicode = unicode;
                while (endIndex + 1 < subset.Tokens.Count)
                {
                    var nextToken = subset.Tokens[endIndex + 1];
                    var nextCode = subset.TokenCodes[nextToken];
                    var nextUnicode = SingleUnicodeScalar(nextToken);
                    if (nextUnicode < 0)
                    {
                        break;
                    }

                    if (nextCode != endCode + 1 || nextUnicode != endUnicode + 1)
                    {
                        break;
                    }

                    endIndex++;
                    endCode = nextCode;
                    endUnicode = nextUnicode;
                }

                if (endIndex > index)
                {
                    ranges.Add(new PdfToUnicodeRange(code, endCode, unicode));
                    index = endIndex + 1;
                    continue;
                }

                singles.Add(new PdfToUnicodeSingle(code, token));
                index++;
            }
        }

        private static int SingleUnicodeScalar(string token)
        {
            if (string.IsNullOrEmpty(token))
            {
                return -1;
            }

            if (token.Length == 1)
            {
                return token[0];
            }

            if (token.Length == 2 && char.IsHighSurrogate(token[0]) && char.IsLowSurrogate(token[1]))
            {
                return char.ConvertToUtf32(token[0], token[1]);
            }

            return -1;
        }

        private static string ToHexUnicode(string text)
        {
            var bytes = Encoding.BigEndianUnicode.GetBytes(text);
            var builder = new StringBuilder(bytes.Length * 2);
            for (var i = 0; i < bytes.Length; i++)
            {
                builder.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

        private static string BuildFontMatrix(bool usePositiveYFontMatrix)
        {
            if (usePositiveYFontMatrix)
            {
                return "[0.001 0 0 0.001 0 0]";
            }

            return "[0.001 0 0 -0.001 0 0]";
        }

        private static bool UsesPositiveYMatrix(SKTypeface typeface)
        {
            if (typeface == null || string.IsNullOrEmpty(typeface.FamilyName))
            {
                return false;
            }

            return string.Equals(typeface.FamilyName, "DengXian", System.StringComparison.OrdinalIgnoreCase)
                || string.Equals(typeface.FamilyName, "绛夌嚎", System.StringComparison.Ordinal);
        }

        private static byte[] BuildStreamObject(int objectId, byte[] streamBytes)
        {
            var compressedObject = BuildCompressedStreamObject(objectId, streamBytes);
            var flatObject = BuildFlatStreamObject(objectId, streamBytes);
            if (flatObject.Length < compressedObject.Length)
            {
                return flatObject;
            }

            return compressedObject;
        }

        private static byte[] BuildCompressedStreamObject(int objectId, byte[] streamBytes)
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

        private static byte[] BuildFlatStreamObject(int objectId, byte[] streamBytes)
        {
            var header = objectId.ToString(CultureInfo.InvariantCulture)
                + " 0 obj\n<< /Length "
                + streamBytes.Length.ToString(CultureInfo.InvariantCulture)
                + " >>\nstream\n";
            var footer = "\nendstream\nendobj\n";
            using (var output = new MemoryStream())
            {
                var headerBytes = Encoding.ASCII.GetBytes(header);
                var footerBytes = Encoding.ASCII.GetBytes(footer);
                output.Write(headerBytes, 0, headerBytes.Length);
                output.Write(streamBytes, 0, streamBytes.Length);
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

        private struct PdfToUnicodeRange
        {
            public readonly byte StartCode;
            public readonly byte EndCode;
            public readonly int StartUnicode;

            public PdfToUnicodeRange(byte startCode, byte endCode, int startUnicode)
            {
                StartCode = startCode;
                EndCode = endCode;
                StartUnicode = startUnicode;
            }
        }

        private struct PdfToUnicodeSingle
        {
            public readonly byte Code;
            public readonly string Token;

            public PdfToUnicodeSingle(byte code, string token)
            {
                Code = code;
                Token = token;
            }
        }
    }
}
