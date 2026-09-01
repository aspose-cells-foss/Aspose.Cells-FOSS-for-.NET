using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfImageObjectWriter
    {
        public byte[] Rewrite(byte[] originalPdf, WorkbookModel workbook)
        {
            if (originalPdf == null || workbook == null)
            {
                return originalPdf;
            }

            var replacements = CollectLosslessPictureStreams(workbook);
            if (replacements.Count == 0)
            {
                return originalPdf;
            }

            var table = PdfObjectTable.Parse(originalPdf);
            var changed = false;
            foreach (var key in new List<int>(table.Objects.Keys))
            {
                var objectBytes = table.Objects[key];
                ImageObjectInfo info;
                if (!TryParseImageObject(objectBytes, out info))
                {
                    continue;
                }

                PictureStreamReplacement replacement;
                if (!TryMatchReplacement(replacements, info.Width, info.Height, out replacement))
                {
                    continue;
                }

                table.Objects[key] = BuildImageObject(key, info.Width, info.Height, replacement.StreamBytes);
                changed = true;
            }

            if (!changed)
            {
                return originalPdf;
            }

            return RebuildPdf(table);
        }

        private static bool TryMatchReplacement(List<PictureStreamReplacement> replacements, int width, int height, out PictureStreamReplacement replacement)
        {
            for (var i = 0; i < replacements.Count; i++)
            {
                if (replacements[i].Used)
                {
                    continue;
                }

                if (replacements[i].Width == width && replacements[i].Height == height)
                {
                    replacements[i].Used = true;
                    replacement = replacements[i];
                    return true;
                }
            }

            replacement = null;
            return false;
        }

        private static List<PictureStreamReplacement> CollectLosslessPictureStreams(WorkbookModel workbook)
        {
            var replacements = new List<PictureStreamReplacement>();
            if (workbook == null || workbook.Worksheets == null)
            {
                return replacements;
            }

            for (var sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                var sheet = workbook.Worksheets[sheetIndex];
                if (sheet == null || sheet.Pictures == null)
                {
                    continue;
                }

                for (var pictureIndex = 0; pictureIndex < sheet.Pictures.Count; pictureIndex++)
                {
                    var picture = sheet.Pictures[pictureIndex];
                    if (!ShouldPreserveLosslessly(picture))
                    {
                        continue;
                    }

                    var replacement = BuildReplacement(picture);
                    if (replacement != null)
                    {
                        replacements.Add(replacement);
                    }
                }
            }

            return replacements;
        }

        private static bool ShouldPreserveLosslessly(PictureModel picture)
        {
            if (picture == null || picture.ImageData == null || picture.ImageData.Length == 0 || string.IsNullOrEmpty(picture.ImageExtension))
            {
                return false;
            }

            var extension = picture.ImageExtension.Trim().ToLowerInvariant();
            return extension == "png" || extension == "gif" || extension == "bmp";
        }

        private static PictureStreamReplacement BuildReplacement(PictureModel picture)
        {
            using (var bitmap = SKBitmap.Decode(picture.ImageData))
            {
                if (bitmap == null || bitmap.Width <= 0 || bitmap.Height <= 0)
                {
                    return null;
                }

                var raw = new byte[bitmap.Width * bitmap.Height * 3];
                var cursor = 0;
                for (var y = 0; y < bitmap.Height; y++)
                {
                    for (var x = 0; x < bitmap.Width; x++)
                    {
                        var color = bitmap.GetPixel(x, y);
                        if (color.Alpha < 255)
                        {
                            var alpha = color.Alpha / 255f;
                            raw[cursor++] = BlendOnWhite(color.Red, alpha);
                            raw[cursor++] = BlendOnWhite(color.Green, alpha);
                            raw[cursor++] = BlendOnWhite(color.Blue, alpha);
                        }
                        else
                        {
                            raw[cursor++] = color.Red;
                            raw[cursor++] = color.Green;
                            raw[cursor++] = color.Blue;
                        }
                    }
                }

                return new PictureStreamReplacement(bitmap.Width, bitmap.Height, Deflate(raw));
            }
        }

        private static byte BlendOnWhite(byte channel, float alpha)
        {
            var value = channel * alpha + 255f * (1f - alpha);
            if (value < 0f)
            {
                value = 0f;
            }

            if (value > 255f)
            {
                value = 255f;
            }

            return (byte)Math.Round(value);
        }

        private static byte[] Deflate(byte[] raw)
        {
            var zlib = TryZlibCompress(raw);
            if (zlib != null)
            {
                return zlib;
            }

            using (var output = new MemoryStream())
            {
                using (var zip = new DeflateStream(output, CompressionLevel.Optimal, true))
                {
                    zip.Write(raw, 0, raw.Length);
                }

                var deflate = output.ToArray();
                return WrapZlib(deflate, raw);
            }
        }

        private static byte[] TryZlibCompress(byte[] raw)
        {
            var zlibStreamType = Type.GetType("System.IO.Compression.ZLibStream, System.IO.Compression", false);
            if (zlibStreamType == null)
            {
                return null;
            }

            var constructor = zlibStreamType.GetConstructor(new[] { typeof(Stream), typeof(CompressionLevel), typeof(bool) });
            if (constructor == null)
            {
                return null;
            }

            using (var output = new MemoryStream())
            {
                IDisposable zip = null;
                try
                {
                    zip = (IDisposable)constructor.Invoke(new object[] { output, CompressionLevel.Optimal, true });
                    var stream = zip as Stream;
                    if (stream == null)
                    {
                        return null;
                    }

                    stream.Write(raw, 0, raw.Length);
                }
                finally
                {
                    if (zip != null)
                    {
                        zip.Dispose();
                    }
                }

                return output.ToArray();
            }
        }

        private static byte[] WrapZlib(byte[] deflate, byte[] raw)
        {
            using (var output = new MemoryStream())
            {
                output.WriteByte(0x78);
                output.WriteByte(0xDA);
                output.Write(deflate, 0, deflate.Length);

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

        private static bool TryParseImageObject(byte[] objectBytes, out ImageObjectInfo info)
        {
            var text = Latin1(objectBytes);
            if (text.IndexOf("/Subtype /Image", StringComparison.Ordinal) < 0)
            {
                info = default(ImageObjectInfo);
                return false;
            }

            var width = ParseIntValue(text, "/Width");
            var height = ParseIntValue(text, "/Height");
            if (width <= 0 || height <= 0)
            {
                info = default(ImageObjectInfo);
                return false;
            }

            info = new ImageObjectInfo(width, height);
            return true;
        }

        private static int ParseIntValue(string text, string key)
        {
            var index = text.IndexOf(key, StringComparison.Ordinal);
            if (index < 0)
            {
                return 0;
            }

            index += key.Length;
            while (index < text.Length && char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            var end = index;
            while (end < text.Length && char.IsDigit(text[end]))
            {
                end++;
            }

            if (end <= index)
            {
                return 0;
            }

            return int.Parse(text.Substring(index, end - index), CultureInfo.InvariantCulture);
        }

        private static byte[] BuildImageObject(int objectId, int width, int height, byte[] streamBytes)
        {
            var header = objectId.ToString(CultureInfo.InvariantCulture)
                + " 0 obj\n<< /Type /XObject\n/Subtype /Image\n/Width "
                + width.ToString(CultureInfo.InvariantCulture)
                + "\n/Height "
                + height.ToString(CultureInfo.InvariantCulture)
                + "\n/ColorSpace /DeviceRGB\n/BitsPerComponent 8\n/Interpolate false\n/Filter /FlateDecode\n/Length "
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

        private static byte[] RebuildPdf(PdfObjectTable table)
        {
            var ids = new List<int>();
            foreach (var key in table.Objects.Keys)
            {
                ids.Add(key);
            }

            ids.Sort();

            using (var output = new MemoryStream())
            {
                output.Write(table.HeaderBytes, 0, table.HeaderBytes.Length);
                var offsets = new Dictionary<int, int>();
                for (var i = 0; i < ids.Count; i++)
                {
                    offsets[ids[i]] = (int)output.Position;
                    var bytes = table.Objects[ids[i]];
                    output.Write(bytes, 0, bytes.Length);
                    if (bytes.Length == 0 || bytes[bytes.Length - 1] != (byte)'\n')
                    {
                        output.WriteByte((byte)'\n');
                    }
                }

                var xrefOffset = (int)output.Position;
                var maxId = ids[ids.Count - 1];
                var xref = new StringBuilder();
                xref.Append("xref\n0 ");
                xref.Append((maxId + 1).ToString(CultureInfo.InvariantCulture));
                xref.Append("\n0000000000 65535 f \n");
                for (var id = 1; id <= maxId; id++)
                {
                    int objectOffset;
                    if (!offsets.TryGetValue(id, out objectOffset))
                    {
                        xref.Append("0000000000 00000 f \n");
                    }
                    else
                    {
                        xref.Append(objectOffset.ToString("D10", CultureInfo.InvariantCulture));
                        xref.Append(" 00000 n \n");
                    }
                }

                xref.Append("trailer\n<< /Size ");
                xref.Append((maxId + 1).ToString(CultureInfo.InvariantCulture));
                if (!string.IsNullOrEmpty(table.RootReference))
                {
                    xref.Append("\n/Root ");
                    xref.Append(table.RootReference);
                }

                if (!string.IsNullOrEmpty(table.InfoReference))
                {
                    xref.Append("\n/Info ");
                    xref.Append(table.InfoReference);
                }

                xref.Append(" >>\nstartxref\n");
                xref.Append(xrefOffset.ToString(CultureInfo.InvariantCulture));
                xref.Append("\n%%EOF");
                var trailerBytes = Encoding.ASCII.GetBytes(xref.ToString());
                output.Write(trailerBytes, 0, trailerBytes.Length);
                return output.ToArray();
            }
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

        private struct ImageObjectInfo
        {
            public readonly int Width;
            public readonly int Height;

            public ImageObjectInfo(int width, int height)
            {
                Width = width;
                Height = height;
            }
        }

        private sealed class PictureStreamReplacement
        {
            public readonly int Width;
            public readonly int Height;
            public readonly byte[] StreamBytes;
            public bool Used;

            public PictureStreamReplacement(int width, int height, byte[] streamBytes)
            {
                Width = width;
                Height = height;
                StreamBytes = streamBytes;
            }
        }
    }
}
