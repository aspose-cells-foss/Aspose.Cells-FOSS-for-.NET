using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfType3TextFontBackend : PdfTextFontBackend
    {
        private readonly PdfFontUsageRegistry _registry;

        public PdfType3TextFontBackend()
        {
            _registry = new PdfFontUsageRegistry();
        }

        public void Initialize(IList<PdfTextRunRecord> runs)
        {
            _registry.AddRuns(runs);
            _registry.FinalizeSubsets();
        }

        public IDictionary<int, byte[]> BuildObjects(ref int nextObjectId)
        {
            return new PdfType3FontBuilder().BuildObjects(_registry, ref nextObjectId);
        }

        public string BuildFontResourceText()
        {
            var builder = new StringBuilder();
            builder.Append("/Font <<");
            for (var i = 0; i < _registry.Usages.Count; i++)
            {
                var usage = _registry.Usages[i];
                for (var s = 0; s < usage.Subsets.Count; s++)
                {
                    var subset = usage.Subsets[s];
                    builder.Append(" /");
                    builder.Append(subset.ResourceName);
                    builder.Append(' ');
                    builder.Append(subset.FontObjectId.ToString(CultureInfo.InvariantCulture));
                    builder.Append(" 0 R");
                }
            }

            builder.Append(" >>");
            return builder.ToString();
        }

        public IList<PdfEncodedTextSegment> EncodeRun(PdfTextRunRecord run)
        {
            var segments = new List<PdfEncodedTextSegment>();
            if (run == null || run.Typeface == null || string.IsNullOrEmpty(run.Text))
            {
                return segments;
            }

            var usage = _registry.UsageFor(run.Typeface, run.UsePositiveYFontMatrix);
            var cursor = run.XPt;
            var index = 0;
            while (index < run.Text.Length)
            {
                var subset = default(PdfType3FontSubset);
                var bytes = new List<byte>();
                var segmentStart = cursor;

                while (index < run.Text.Length)
                {
                    var token = ReadToken(run.Text, ref index);
                    PdfType3FontSubset tokenSubset;
                    byte code;
                    float width1000;
                    if (!usage.TryResolveToken(token, out tokenSubset, out code, out width1000))
                    {
                        continue;
                    }

                    if (subset == null)
                    {
                        subset = tokenSubset;
                    }
                    else if (!ReferenceEquals(subset, tokenSubset))
                    {
                        index -= token.Length;
                        break;
                    }

                    bytes.Add(code);
                    cursor += run.FontSizePt * (width1000 / 1000f);
                }

                if (subset != null && bytes.Count > 0)
                {
                    var segment = new PdfEncodedTextSegment();
                    segment.ResourceName = subset.ResourceName;
                    segment.EncodedBytes = bytes.ToArray();
                    segment.StartXPt = segmentStart;
                    segments.Add(segment);
                }
            }

            return segments;
        }

        private static string ReadToken(string text, ref int index)
        {
            var ch = text[index];
            if (char.IsHighSurrogate(ch) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]))
            {
                var token = text.Substring(index, 2);
                index += 2;
                return token;
            }

            index++;
            return text.Substring(index - 1, 1);
        }

        public void Dispose()
        {
            _registry.Dispose();
        }
    }
}
