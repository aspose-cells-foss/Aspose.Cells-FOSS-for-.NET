using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfTextObjectWriter
    {
        private readonly bool _useTrueTypeSubsetPrototype;

        public PdfTextObjectWriter()
            : this(false)
        {
        }

        public PdfTextObjectWriter(bool useTrueTypeSubsetPrototype)
        {
            _useTrueTypeSubsetPrototype = useTrueTypeSubsetPrototype;
        }

        public byte[] Rewrite(byte[] originalPdf, PdfTextDocumentSession session)
        {
            if (originalPdf == null || session == null || !session.HasRuns)
            {
                return originalPdf;
            }

            using (var backend = CreateFontBackend())
            {
                backend.Initialize(session.Runs);

                var table = PdfObjectTable.Parse(originalPdf);
                var nextObjectId = table.MaxObjectId + 1;
                var fontObjects = backend.BuildObjects(ref nextObjectId);
                foreach (var pair in fontObjects)
                {
                    table.Objects[pair.Key] = pair.Value;
                }

                var fontResourceText = backend.BuildFontResourceText();
                for (var pageIndex = 0; pageIndex < table.PageObjectIds.Count; pageIndex++)
                {
                    var pageRuns = RunsForPage(session.Runs, pageIndex + 1);
                    if (pageRuns.Count == 0)
                    {
                        continue;
                    }

                    var pageObjectId = table.PageObjectIds[pageIndex];
                    var pageObjectText = Latin1(table.Objects[pageObjectId]);
                    var pageHeight = ParsePageHeight(pageObjectText);
                    var prologueObjectId = nextObjectId++;
                    var contentObjectId = nextObjectId++;
                    table.Objects[prologueObjectId] = BuildStreamObject(prologueObjectId, Encoding.ASCII.GetBytes("q\n"));
                    table.Objects[contentObjectId] = BuildStreamObject(contentObjectId, BuildPageTextContent(pageRuns, backend, pageHeight));

                    pageObjectText = AddFontResources(pageObjectText, fontResourceText);
                    pageObjectText = AppendContentReferences(pageObjectText, prologueObjectId, contentObjectId);
                    table.Objects[pageObjectId] = Encoding.ASCII.GetBytes(pageObjectText);
                }

                return RebuildPdf(table);
            }
        }

        private PdfTextFontBackend CreateFontBackend()
        {
            if (_useTrueTypeSubsetPrototype)
            {
                return new PdfTrueTypePrototypeBackend();
            }

            return new PdfType3TextFontBackend();
        }

        private static List<PdfTextRunRecord> RunsForPage(IList<PdfTextRunRecord> runs, int pageNumber)
        {
            var pageRuns = new List<PdfTextRunRecord>();
            for (var i = 0; i < runs.Count; i++)
            {
                if (runs[i].PageNumber == pageNumber)
                {
                    pageRuns.Add(runs[i]);
                }
            }

            return pageRuns;
        }

        private static string AddFontResources(string pageObjectText, string fontResources)
        {
            if (pageObjectText.IndexOf("/Font <<", StringComparison.Ordinal) >= 0)
            {
                return pageObjectText;
            }

            var marker = "/Resources <<";
            var index = pageObjectText.IndexOf(marker, StringComparison.Ordinal);
            if (index < 0)
            {
                return pageObjectText;
            }

            return pageObjectText.Insert(index + marker.Length, fontResources);
        }

        private static string AppendContentReferences(string pageObjectText, int prologueObjectId, int contentObjectId)
        {
            var marker = "/Contents ";
            var index = pageObjectText.IndexOf(marker, StringComparison.Ordinal);
            if (index < 0)
            {
                return pageObjectText;
            }

            var start = index + marker.Length;
            if (pageObjectText[start] == '[')
            {
                var endArray = pageObjectText.IndexOf(']', start);
                if (endArray < 0)
                {
                    return pageObjectText;
                }

                return pageObjectText.Insert(
                    endArray,
                    " "
                    + prologueObjectId.ToString(CultureInfo.InvariantCulture)
                    + " 0 R "
                    + contentObjectId.ToString(CultureInfo.InvariantCulture)
                    + " 0 R");
            }

            var end = pageObjectText.IndexOf('\n', start);
            if (end < 0)
            {
                end = pageObjectText.Length;
            }

            var existing = pageObjectText.Substring(start, end - start).Trim();
            var replacement = "["
                + prologueObjectId.ToString(CultureInfo.InvariantCulture)
                + " 0 R "
                + existing
                + " "
                + contentObjectId.ToString(CultureInfo.InvariantCulture)
                + " 0 R]";
            return pageObjectText.Substring(0, start) + replacement + pageObjectText.Substring(end);
        }

        private static byte[] BuildPageTextContent(IList<PdfTextRunRecord> runs, PdfTextFontBackend backend, float pageHeight)
        {
            var builder = new StringBuilder();
            builder.Append("Q\n");

            var start = 0;
            while (start < runs.Count)
            {
                var end = start;
                while (end + 1 < runs.Count && CanBatchRuns(runs[end], runs[end + 1]))
                {
                    end++;
                }

                AppendRunBatch(builder, runs, start, end, backend);
                start = end + 1;
            }

            return Encoding.ASCII.GetBytes(builder.ToString());
        }

        private static void AppendRunBatch(StringBuilder builder, IList<PdfTextRunRecord> runs, int start, int end, PdfTextFontBackend backend)
        {
            var firstRun = runs[start];
            builder.Append("q\n");
            if (firstRun.UseTopDownCoordinates)
            {
                builder.Append("1 0 0 -1 0 ");
                builder.Append(Format(firstRun.PageHeightPt));
                builder.Append(" cm\n");
                builder.Append(Format(firstRun.ClipLeftPt));
                builder.Append(' ');
                builder.Append(Format(firstRun.ClipTopPt));
                builder.Append(' ');
                builder.Append(Format(firstRun.ClipWidthPt));
                builder.Append(' ');
                builder.Append(Format(firstRun.ClipHeightPt));
                builder.Append(" re W n\n");
            }
            else
            {
                builder.Append(Format(firstRun.ClipLeftPt));
                builder.Append(' ');
                builder.Append(Format(firstRun.ClipBottomPt));
                builder.Append(' ');
                builder.Append(Format(firstRun.ClipWidthPt));
                builder.Append(' ');
                builder.Append(Format(firstRun.ClipHeightPt));
                builder.Append(" re W n\n");
            }

            builder.Append("BT\n");
            builder.Append(ColorCommand(firstRun));
            builder.Append('\n');

            string activeResourceName = null;
            float activeFontSize = -1f;

            for (var i = start; i <= end; i++)
            {
                var run = runs[i];
                var segments = backend.EncodeRun(run);
                for (var segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++)
                {
                    var segment = segments[segmentIndex];
                    if (segment == null || segment.EncodedBytes == null || segment.EncodedBytes.Length == 0)
                    {
                        continue;
                    }

                    if (!string.Equals(activeResourceName, segment.ResourceName, StringComparison.Ordinal)
                        || Math.Abs(activeFontSize - run.FontSizePt) > 0.001f)
                    {
                        builder.Append("/");
                        builder.Append(segment.ResourceName);
                        builder.Append(' ');
                        builder.Append(Format(run.FontSizePt));
                        builder.Append(" Tf\n");
                        activeResourceName = segment.ResourceName;
                        activeFontSize = run.FontSizePt;
                    }

                    var hasRotation = Math.Abs(run.RotationDeg) > 0.01f;
                    if (hasRotation)
                    {
                        var radians = run.RotationDeg * (float)Math.PI / 180f;
                        var cos = (float)Math.Cos(radians);
                        var sin = (float)Math.Sin(radians);
                        var e = run.TransformOriginXPt + cos * segment.StartXPt - sin * run.BaselineYPt;
                        var f = run.TransformOriginYPt + sin * segment.StartXPt + cos * run.BaselineYPt;
                        builder.Append(Format(cos));
                        builder.Append(' ');
                        builder.Append(Format(sin));
                        builder.Append(' ');
                        if (run.UseTopDownCoordinates && segment.RequiresTopDownTextFlip)
                        {
                            builder.Append(Format(sin));
                        }
                        else
                        {
                            builder.Append(Format(-sin));
                        }
                        builder.Append(' ');
                        if (run.UseTopDownCoordinates && segment.RequiresTopDownTextFlip)
                        {
                            builder.Append(Format(-cos));
                        }
                        else
                        {
                            builder.Append(Format(cos));
                        }
                        builder.Append(' ');
                        builder.Append(Format(e));
                        builder.Append(' ');
                        builder.Append(Format(f));
                        builder.Append(" Tm\n<");
                    }
                    else
                    {
                        if (run.UseTopDownCoordinates && segment.RequiresTopDownTextFlip)
                        {
                            builder.Append("1 0 0 -1 ");
                        }
                        else
                        {
                            builder.Append("1 0 0 1 ");
                        }

                        builder.Append(Format(segment.StartXPt));
                        builder.Append(' ');
                        builder.Append(Format(run.BaselineYPt));
                        builder.Append(" Tm\n<");
                    }

                    for (var b = 0; b < segment.EncodedBytes.Length; b++)
                    {
                        builder.Append(segment.EncodedBytes[b].ToString("X2", CultureInfo.InvariantCulture));
                    }

                    builder.Append("> Tj\n");
                }
            }

            builder.Append("ET\nQ\n");
        }

        private static bool CanBatchRuns(PdfTextRunRecord left, PdfTextRunRecord right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            if (left.PageNumber != right.PageNumber)
            {
                return false;
            }

            if (left.UseTopDownCoordinates != right.UseTopDownCoordinates)
            {
                return false;
            }

            if (Math.Abs(left.PageHeightPt - right.PageHeightPt) > 0.01f)
            {
                return false;
            }

            if (!SameFloat(left.ClipLeftPt, right.ClipLeftPt)
                || !SameFloat(left.ClipTopPt, right.ClipTopPt)
                || !SameFloat(left.ClipBottomPt, right.ClipBottomPt)
                || !SameFloat(left.ClipWidthPt, right.ClipWidthPt)
                || !SameFloat(left.ClipHeightPt, right.ClipHeightPt))
            {
                return false;
            }

            if (left.Color.Red != right.Color.Red
                || left.Color.Green != right.Color.Green
                || left.Color.Blue != right.Color.Blue)
            {
                return false;
            }

            return true;
        }

        private static bool SameFloat(float left, float right)
        {
            return Math.Abs(left - right) <= 0.01f;
        }

        private static float ParsePageHeight(string pageObjectText)
        {
            var marker = "/MediaBox [";
            var index = pageObjectText.IndexOf(marker, StringComparison.Ordinal);
            if (index < 0)
            {
                return 842f;
            }

            index += marker.Length;
            var end = pageObjectText.IndexOf(']', index);
            if (end < 0)
            {
                return 842f;
            }

            var parts = pageObjectText.Substring(index, end - index)
                .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 4)
            {
                return 842f;
            }

            float y0;
            float y1;
            if (!float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out y0))
            {
                return 842f;
            }

            if (!float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out y1))
            {
                return 842f;
            }

            return Math.Abs(y1 - y0);
        }

        private static string ColorCommand(PdfTextRunRecord run)
        {
            return Format(run.Color.Red / 255f)
                + " "
                + Format(run.Color.Green / 255f)
                + " "
                + Format(run.Color.Blue / 255f)
                + " rg";
        }

        private static byte[] BuildStreamObject(int objectId, byte[] streamBytes)
        {
            return BuildFlatStreamObject(objectId, streamBytes);
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

        private static string Format(float value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
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
