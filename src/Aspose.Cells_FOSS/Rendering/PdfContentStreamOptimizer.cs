using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfContentStreamOptimizer
    {
        private static readonly Regex DecimalTokenPattern = new Regex(@"(?<![A-Za-z0-9_/])[-+]?(?:\d+\.\d+|\.\d+)", RegexOptions.Compiled);

        public byte[] Rewrite(byte[] originalPdf)
        {
            if (originalPdf == null || originalPdf.Length == 0)
            {
                return originalPdf;
            }

            var table = PdfObjectTable.Parse(originalPdf);
            var contentObjectIds = CollectPageContentObjectIds(table);
            if (contentObjectIds.Count == 0)
            {
                return originalPdf;
            }

            var changed = false;
            foreach (var objectId in contentObjectIds)
            {
                byte[] objectBytes;
                if (!table.Objects.TryGetValue(objectId, out objectBytes))
                {
                    continue;
                }

                byte[] rewritten;
                if (!TryOptimizeStreamObject(objectId, objectBytes, out rewritten))
                {
                    continue;
                }

                table.Objects[objectId] = rewritten;
                changed = true;
            }

            if (!changed)
            {
                return originalPdf;
            }

            return RebuildPdf(table);
        }

        private static List<int> CollectPageContentObjectIds(PdfObjectTable table)
        {
            var result = new List<int>();
            var seen = new HashSet<int>();
            for (var i = 0; i < table.PageObjectIds.Count; i++)
            {
                byte[] objectBytes;
                if (!table.Objects.TryGetValue(table.PageObjectIds[i], out objectBytes))
                {
                    continue;
                }

                var pageText = Latin1(objectBytes);
                var contentIds = ParseContentObjectIds(pageText);
                for (var c = 0; c < contentIds.Count; c++)
                {
                    if (seen.Add(contentIds[c]))
                    {
                        result.Add(contentIds[c]);
                    }
                }
            }

            return result;
        }

        private static List<int> ParseContentObjectIds(string pageObjectText)
        {
            var result = new List<int>();
            var marker = "/Contents ";
            var index = pageObjectText.IndexOf(marker, StringComparison.Ordinal);
            if (index < 0)
            {
                return result;
            }

            index += marker.Length;
            if (index >= pageObjectText.Length)
            {
                return result;
            }

            if (pageObjectText[index] == '[')
            {
                var endArray = pageObjectText.IndexOf(']', index);
                if (endArray < 0)
                {
                    return result;
                }

                var items = pageObjectText.Substring(index + 1, endArray - index - 1)
                    .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (var i = 0; i + 2 < items.Length; i++)
                {
                    int objectId;
                    if (int.TryParse(items[i], NumberStyles.Integer, CultureInfo.InvariantCulture, out objectId)
                        && items[i + 1] == "0"
                        && items[i + 2] == "R")
                    {
                        result.Add(objectId);
                    }
                }

                return result;
            }

            var end = pageObjectText.IndexOf('\n', index);
            if (end < 0)
            {
                end = pageObjectText.Length;
            }

            var single = pageObjectText.Substring(index, end - index).Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (single.Length >= 3)
            {
                int objectId;
                if (int.TryParse(single[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out objectId)
                    && single[1] == "0"
                    && single[2] == "R")
                {
                    result.Add(objectId);
                }
            }

            return result;
        }

        private static bool TryOptimizeStreamObject(int objectId, byte[] objectBytes, out byte[] rewritten)
        {
            rewritten = null;
            var objectText = Latin1(objectBytes);
            if (objectText.IndexOf("/Filter /FlateDecode", StringComparison.Ordinal) < 0)
            {
                return false;
            }

            var streamMarker = "stream\n";
            var streamIndex = objectText.IndexOf(streamMarker, StringComparison.Ordinal);
            if (streamIndex < 0)
            {
                streamMarker = "stream\r\n";
                streamIndex = objectText.IndexOf(streamMarker, StringComparison.Ordinal);
                if (streamIndex < 0)
                {
                    return false;
                }
            }

            var streamStart = streamIndex + streamMarker.Length;
            var endStreamIndex = objectText.IndexOf("\nendstream", streamStart, StringComparison.Ordinal);
            var endStreamMarkerLength = 1;
            if (endStreamIndex < 0)
            {
                endStreamIndex = objectText.IndexOf("\r\nendstream", streamStart, StringComparison.Ordinal);
                endStreamMarkerLength = 2;
                if (endStreamIndex < 0)
                {
                    return false;
                }
            }

            var compressed = new byte[endStreamIndex - streamStart];
            Buffer.BlockCopy(objectBytes, streamStart, compressed, 0, compressed.Length);

            byte[] decompressed;
            try
            {
                decompressed = Decompress(compressed);
            }
            catch (InvalidDataException)
            {
                return false;
            }

            var optimized = OptimizeContentText(Encoding.ASCII.GetString(decompressed));
            var optimizedBytes = Encoding.ASCII.GetBytes(optimized);
            if (optimizedBytes.Length >= decompressed.Length)
            {
                return false;
            }

            var recompressed = Compress(optimizedBytes);
            var header = objectText.Substring(0, streamIndex);
            header = ReplaceLength(header, recompressed.Length);
            var footer = objectText.Substring(endStreamIndex + endStreamMarkerLength);

            using (var output = new MemoryStream())
            {
                var headerBytes = Encoding.ASCII.GetBytes(header + streamMarker);
                var footerBytes = Encoding.ASCII.GetBytes(footer);
                output.Write(headerBytes, 0, headerBytes.Length);
                output.Write(recompressed, 0, recompressed.Length);
                output.Write(footerBytes, 0, footerBytes.Length);
                rewritten = output.ToArray();
            }

            return true;
        }

        private static string ReplaceLength(string header, int newLength)
        {
            var marker = "/Length ";
            var index = header.IndexOf(marker, StringComparison.Ordinal);
            if (index < 0)
            {
                return header;
            }

            var start = index + marker.Length;
            var end = start;
            while (end < header.Length && char.IsDigit(header[end]))
            {
                end++;
            }

            return header.Substring(0, start)
                + newLength.ToString(CultureInfo.InvariantCulture)
                + header.Substring(end);
        }

        private static string OptimizeContentText(string content)
        {
            var compact = DecimalTokenPattern.Replace(content, delegate (Match match)
            {
                double value;
                if (!double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    return match.Value;
                }

                return FormatCompact(value);
            });

            compact = CompactWhitespace(compact);
            return RemoveRedundantGraphicsStateCommands(compact);
        }

        private static string RemoveRedundantGraphicsStateCommands(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }

            var tokens = TokenizeContent(content);
            if (tokens.Count == 0)
            {
                return content;
            }

            var output = new List<string>(tokens.Count);
            var operands = new List<string>();
            string strokeRgb = null;
            string fillRgb = null;
            string strokeGray = null;
            string fillGray = null;
            string graphicsState = null;
            var strokeRgbStack = new List<string>();
            var fillRgbStack = new List<string>();
            var strokeGrayStack = new List<string>();
            var fillGrayStack = new List<string>();
            var graphicsStateStack = new List<string>();

            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (!IsOperatorToken(token))
                {
                    operands.Add(token);
                    continue;
                }

                if (token == "q" && operands.Count == 0)
                {
                    output.Add(token);
                    strokeRgbStack.Add(strokeRgb);
                    fillRgbStack.Add(fillRgb);
                    strokeGrayStack.Add(strokeGray);
                    fillGrayStack.Add(fillGray);
                    graphicsStateStack.Add(graphicsState);
                    continue;
                }

                if (token == "Q" && operands.Count == 0)
                {
                    output.Add(token);
                    var lastIndex = strokeRgbStack.Count - 1;
                    if (lastIndex >= 0)
                    {
                        strokeRgb = strokeRgbStack[lastIndex];
                        fillRgb = fillRgbStack[lastIndex];
                        strokeGray = strokeGrayStack[lastIndex];
                        fillGray = fillGrayStack[lastIndex];
                        graphicsState = graphicsStateStack[lastIndex];
                        strokeRgbStack.RemoveAt(lastIndex);
                        fillRgbStack.RemoveAt(lastIndex);
                        strokeGrayStack.RemoveAt(lastIndex);
                        fillGrayStack.RemoveAt(lastIndex);
                        graphicsStateStack.RemoveAt(lastIndex);
                    }
                    else
                    {
                        strokeRgb = null;
                        fillRgb = null;
                        strokeGray = null;
                        fillGray = null;
                        graphicsState = null;
                    }

                    continue;
                }

                if (token == "RG" && operands.Count == 3)
                {
                    var value = JoinOperands(operands);
                    if (!string.Equals(strokeRgb, value, StringComparison.Ordinal))
                    {
                        FlushOperation(output, operands, token);
                        strokeRgb = value;
                    }
                    else
                    {
                        operands.Clear();
                    }

                    continue;
                }

                if (token == "rg" && operands.Count == 3)
                {
                    var value = JoinOperands(operands);
                    if (!string.Equals(fillRgb, value, StringComparison.Ordinal))
                    {
                        FlushOperation(output, operands, token);
                        fillRgb = value;
                    }
                    else
                    {
                        operands.Clear();
                    }

                    continue;
                }

                if (token == "G" && operands.Count == 1)
                {
                    var value = operands[0];
                    if (!string.Equals(strokeGray, value, StringComparison.Ordinal))
                    {
                        FlushOperation(output, operands, token);
                        strokeGray = value;
                    }
                    else
                    {
                        operands.Clear();
                    }

                    continue;
                }

                if (token == "g" && operands.Count == 1)
                {
                    var value = operands[0];
                    if (!string.Equals(fillGray, value, StringComparison.Ordinal))
                    {
                        FlushOperation(output, operands, token);
                        fillGray = value;
                    }
                    else
                    {
                        operands.Clear();
                    }

                    continue;
                }

                if (token == "gs" && operands.Count == 1)
                {
                    var value = operands[0];
                    if (!string.Equals(graphicsState, value, StringComparison.Ordinal))
                    {
                        FlushOperation(output, operands, token);
                        graphicsState = value;
                    }
                    else
                    {
                        operands.Clear();
                    }

                    continue;
                }

                FlushOperation(output, operands, token);
            }

            if (operands.Count > 0)
            {
                for (var i = 0; i < operands.Count; i++)
                {
                    output.Add(operands[i]);
                }
            }

            return string.Join(" ", output);
        }

        private static List<string> TokenizeContent(string content)
        {
            var tokens = new List<string>();
            var builder = new StringBuilder();
            var inLiteralString = false;
            var inHexString = false;
            var literalDepth = 0;
            var escaping = false;

            for (var i = 0; i < content.Length; i++)
            {
                var ch = content[i];
                if (inLiteralString)
                {
                    builder.Append(ch);
                    if (escaping)
                    {
                        escaping = false;
                    }
                    else if (ch == '\\')
                    {
                        escaping = true;
                    }
                    else if (ch == '(')
                    {
                        literalDepth++;
                    }
                    else if (ch == ')')
                    {
                        literalDepth--;
                        if (literalDepth <= 0)
                        {
                            inLiteralString = false;
                            AddToken(tokens, builder);
                        }
                    }

                    continue;
                }

                if (inHexString)
                {
                    builder.Append(ch);
                    if (ch == '>')
                    {
                        inHexString = false;
                        AddToken(tokens, builder);
                    }

                    continue;
                }

                if (char.IsWhiteSpace(ch))
                {
                    AddToken(tokens, builder);
                    continue;
                }

                if (ch == '(')
                {
                    AddToken(tokens, builder);
                    builder.Append(ch);
                    inLiteralString = true;
                    literalDepth = 1;
                    escaping = false;
                    continue;
                }

                if (ch == '<' && !IsDictionaryStart(content, i))
                {
                    AddToken(tokens, builder);
                    builder.Append(ch);
                    inHexString = true;
                    continue;
                }

                builder.Append(ch);
            }

            AddToken(tokens, builder);
            return tokens;
        }

        private static void AddToken(List<string> tokens, StringBuilder builder)
        {
            if (builder.Length == 0)
            {
                return;
            }

            tokens.Add(builder.ToString());
            builder.Clear();
        }

        private static bool IsOperatorToken(string token)
        {
            switch (token)
            {
                case "q":
                case "Q":
                case "cm":
                case "w":
                case "J":
                case "j":
                case "M":
                case "d":
                case "ri":
                case "i":
                case "gs":
                case "m":
                case "l":
                case "c":
                case "v":
                case "y":
                case "h":
                case "re":
                case "S":
                case "s":
                case "f":
                case "F":
                case "f*":
                case "B":
                case "B*":
                case "b":
                case "b*":
                case "n":
                case "W":
                case "W*":
                case "BT":
                case "ET":
                case "Tf":
                case "Tm":
                case "Td":
                case "TD":
                case "Tj":
                case "TJ":
                case "Tc":
                case "Tw":
                case "Tz":
                case "TL":
                case "Tr":
                case "Ts":
                case "Do":
                case "RG":
                case "rg":
                case "G":
                case "g":
                case "K":
                case "k":
                case "CS":
                case "cs":
                case "SC":
                case "SCN":
                case "sc":
                case "scn":
                case "sh":
                    return true;
                default:
                    return false;
            }
        }

        private static string JoinOperands(List<string> operands)
        {
            return string.Join(" ", operands);
        }

        private static void FlushOperation(List<string> output, List<string> operands, string operatorToken)
        {
            for (var i = 0; i < operands.Count; i++)
            {
                output.Add(operands[i]);
            }

            operands.Clear();
            output.Add(operatorToken);
        }

        private static string CompactWhitespace(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }

            var builder = new StringBuilder(content.Length);
            var pendingSpace = false;
            var inLiteralString = false;
            var literalDepth = 0;
            var escaping = false;
            var inHexString = false;
            var inComment = false;

            for (var i = 0; i < content.Length; i++)
            {
                var ch = content[i];

                if (inComment)
                {
                    if (ch == '\r' || ch == '\n')
                    {
                        inComment = false;
                        pendingSpace = builder.Length > 0;
                    }

                    continue;
                }

                if (inLiteralString)
                {
                    builder.Append(ch);
                    if (escaping)
                    {
                        escaping = false;
                    }
                    else if (ch == '\\')
                    {
                        escaping = true;
                    }
                    else if (ch == '(')
                    {
                        literalDepth++;
                    }
                    else if (ch == ')')
                    {
                        literalDepth--;
                        if (literalDepth <= 0)
                        {
                            inLiteralString = false;
                            literalDepth = 0;
                        }
                    }

                    continue;
                }

                if (inHexString)
                {
                    builder.Append(ch);
                    if (ch == '>')
                    {
                        inHexString = false;
                    }

                    continue;
                }

                if (char.IsWhiteSpace(ch))
                {
                    pendingSpace = builder.Length > 0;
                    continue;
                }

                if (ch == '%')
                {
                    inComment = true;
                    continue;
                }

                if (pendingSpace && ShouldEmitSpace(builder))
                {
                    builder.Append(' ');
                }

                pendingSpace = false;
                builder.Append(ch);

                if (ch == '(')
                {
                    inLiteralString = true;
                    literalDepth = 1;
                    escaping = false;
                }
                else if (ch == '<' && !IsDictionaryStart(content, i))
                {
                    inHexString = true;
                }
            }

            return builder.ToString();
        }

        private static bool ShouldEmitSpace(StringBuilder builder)
        {
            if (builder.Length == 0)
            {
                return false;
            }

            var previous = builder[builder.Length - 1];
            if (previous == '[' || previous == '<' || previous == '(' || previous == ' ')
            {
                return false;
            }

            return true;
        }

        private static bool IsDictionaryStart(string content, int index)
        {
            return index + 1 < content.Length && content[index + 1] == '<';
        }

        private static string FormatCompact(double value)
        {
            // Keep enough precision for color operators and fine geometry while still shrinking
            // Skia's verbose decimal output. One decimal turns RGB fills like 0.851 0.882 0.949
            // into 0.9 0.9 0.9, which visibly shifts Excel's pale fills toward gray.
            var rounded = Math.Round(value, 3, MidpointRounding.AwayFromZero);
            if (rounded == 0d)
            {
                return "0";
            }

            var text = rounded.ToString("0.###", CultureInfo.InvariantCulture);
            if (text.StartsWith("0.", StringComparison.Ordinal))
            {
                return text.Substring(1);
            }

            if (text.StartsWith("-0.", StringComparison.Ordinal))
            {
                return "-" + text.Substring(2);
            }

            return text;
        }

        private static byte[] Decompress(byte[] compressed)
        {
            using (var input = new MemoryStream(compressed))
            using (var zlib = OpenZlibStream(input))
            using (var output = new MemoryStream())
            {
                zlib.CopyTo(output);
                return output.ToArray();
            }
        }

        private static Stream OpenZlibStream(Stream input)
        {
            var zlibStreamType = Type.GetType("System.IO.Compression.ZLibStream, System.IO.Compression", false);
            if (zlibStreamType != null)
            {
                var constructor = zlibStreamType.GetConstructor(new[] { typeof(Stream), typeof(CompressionMode), typeof(bool) });
                if (constructor != null)
                {
                    return (Stream)constructor.Invoke(new object[] { input, CompressionMode.Decompress, true });
                }
            }

            if (input.ReadByte() < 0 || input.ReadByte() < 0)
            {
                throw new InvalidDataException("Invalid zlib stream.");
            }

            return new DeflateStream(input, CompressionMode.Decompress, true);
        }

        private static byte[] Compress(byte[] raw)
        {
            var zlibStreamType = Type.GetType("System.IO.Compression.ZLibStream, System.IO.Compression", false);
            if (zlibStreamType != null)
            {
                var constructor = zlibStreamType.GetConstructor(new[] { typeof(Stream), typeof(CompressionLevel), typeof(bool) });
                if (constructor != null)
                {
                    using (var output = new MemoryStream())
                    {
                        IDisposable stream = null;
                        try
                        {
                            stream = (IDisposable)constructor.Invoke(new object[] { output, CompressionLevel.Optimal, true });
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
                }

                var xrefStart = (int)output.Position;
                var maxId = ids.Count > 0 ? ids[ids.Count - 1] : 0;
                var xref = new StringBuilder();
                xref.Append("xref\n0 ");
                xref.Append((maxId + 1).ToString(CultureInfo.InvariantCulture));
                xref.Append("\n0000000000 65535 f \n");
                for (var objectId = 1; objectId <= maxId; objectId++)
                {
                    int offset;
                    if (offsets.TryGetValue(objectId, out offset))
                    {
                        xref.Append(offset.ToString("D10", CultureInfo.InvariantCulture));
                        xref.Append(" 00000 n \n");
                    }
                    else
                    {
                        xref.Append("0000000000 00000 f \n");
                    }
                }

                output.Write(Encoding.ASCII.GetBytes(xref.ToString()), 0, xref.Length);

                var trailer = new StringBuilder();
                trailer.Append("trailer\n<< /Size ");
                trailer.Append((maxId + 1).ToString(CultureInfo.InvariantCulture));
                if (!string.IsNullOrEmpty(table.RootReference))
                {
                    trailer.Append(" /Root ");
                    trailer.Append(table.RootReference);
                }

                if (!string.IsNullOrEmpty(table.InfoReference))
                {
                    trailer.Append(" /Info ");
                    trailer.Append(table.InfoReference);
                }

                trailer.Append(" >>\nstartxref\n");
                trailer.Append(xrefStart.ToString(CultureInfo.InvariantCulture));
                trailer.Append("\n%%EOF");
                output.Write(Encoding.ASCII.GetBytes(trailer.ToString()), 0, trailer.Length);

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
    }
}
