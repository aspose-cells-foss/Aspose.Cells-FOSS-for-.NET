using System;
using System.Collections.Generic;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfTrueTypePrototypeBackend : PdfTextFontBackend
    {
        private readonly PdfType3TextFontBackend _fallback;
        private readonly List<SKTypeface> _probedTypefaces;
        private readonly List<PdfTrueTypePrototypeUsage> _trueTypeUsages;

        public PdfTrueTypePrototypeBackend()
        {
            _fallback = new PdfType3TextFontBackend();
            _probedTypefaces = new List<SKTypeface>();
            _trueTypeUsages = new List<PdfTrueTypePrototypeUsage>();
        }

        public void Initialize(IList<PdfTextRunRecord> runs)
        {
            var fallbackRuns = new List<PdfTextRunRecord>();
            if (runs != null)
            {
                for (var i = 0; i < runs.Count; i++)
                {
                    var run = runs[i];
                    if (IsEligibleForPrototype(run))
                    {
                        var usage = UsageFor(run.Typeface);
                        usage.AddText(run.Text);
                    }
                    else
                    {
                        fallbackRuns.Add(run);
                    }
                }
            }

            _fallback.Initialize(fallbackRuns);
        }

        public IDictionary<int, byte[]> BuildObjects(ref int nextObjectId)
        {
            var objects = new Dictionary<int, byte[]>();
            var trueTypeObjects = new PdfTrueTypePrototypeFontBuilder().BuildObjects(_trueTypeUsages, ref nextObjectId);
            foreach (var pair in trueTypeObjects)
            {
                objects[pair.Key] = pair.Value;
            }

            var fallbackObjects = _fallback.BuildObjects(ref nextObjectId);
            foreach (var pair in fallbackObjects)
            {
                objects[pair.Key] = pair.Value;
            }

            return objects;
        }

        public string BuildFontResourceText()
        {
            var builder = new System.Text.StringBuilder();
            builder.Append("/Font <<");
            for (var i = 0; i < _trueTypeUsages.Count; i++)
            {
                var usage = _trueTypeUsages[i];
                builder.Append(" /");
                builder.Append(usage.ResourceName);
                builder.Append(' ');
                builder.Append(usage.FontObjectId.ToString(System.Globalization.CultureInfo.InvariantCulture));
                builder.Append(" 0 R");
            }

            var fallbackText = _fallback.BuildFontResourceText();
            var start = fallbackText.IndexOf("<<", StringComparison.Ordinal);
            var end = fallbackText.LastIndexOf(">>", StringComparison.Ordinal);
            if (start >= 0 && end > start + 2)
            {
                builder.Append(fallbackText.Substring(start + 2, end - start - 2));
            }

            builder.Append(" >>");
            return builder.ToString();
        }

        public IList<PdfEncodedTextSegment> EncodeRun(PdfTextRunRecord run)
        {
            if (!IsEligibleForPrototype(run))
            {
                return _fallback.EncodeRun(run);
            }

            var usage = UsageFor(run.Typeface);
            var segment = new PdfEncodedTextSegment();
            segment.ResourceName = usage.ResourceName;
            segment.EncodedBytes = usage.EncodeText(run.Text);
            segment.StartXPt = run.XPt;
            segment.RequiresTopDownTextFlip = true;
            return new PdfEncodedTextSegment[] { segment };
        }

        private bool IsEligibleForPrototype(PdfTextRunRecord run)
        {
            if (run == null || run.Typeface == null || string.IsNullOrEmpty(run.Text))
            {
                return false;
            }

            if (!run.UseTopDownCoordinates)
            {
                return false;
            }

            for (var i = 0; i < run.Text.Length; i++)
            {
                var ch = run.Text[i];
                if (ch < 32 || ch > 126)
                {
                    return false;
                }
            }

            ValidateRequiredTables(run.Typeface);
            return true;
        }

        private PdfTrueTypePrototypeUsage UsageFor(SKTypeface typeface)
        {
            for (var i = 0; i < _trueTypeUsages.Count; i++)
            {
                if (ReferenceEquals(_trueTypeUsages[i].Typeface, typeface))
                {
                    return _trueTypeUsages[i];
                }
            }

            var usage = new PdfTrueTypePrototypeUsage(typeface);
            _trueTypeUsages.Add(usage);
            return usage;
        }

        private bool AlreadyProbed(SKTypeface typeface)
        {
            for (var i = 0; i < _probedTypefaces.Count; i++)
            {
                if (ReferenceEquals(_probedTypefaces[i], typeface))
                {
                    return true;
                }
            }

            return false;
        }

        private static void ValidateRequiredTables(SKTypeface typeface)
        {
            if (typeface == null)
            {
                throw new InvalidOperationException("TrueType prototype requires a resolved typeface.");
            }

            var requiredTags = new uint[]
            {
                Tag('c', 'm', 'a', 'p'),
                Tag('h', 'e', 'a', 'd'),
                Tag('h', 'h', 'e', 'a'),
                Tag('h', 'm', 't', 'x'),
                Tag('m', 'a', 'x', 'p'),
                Tag('l', 'o', 'c', 'a'),
                Tag('g', 'l', 'y', 'f')
            };

            var availableTags = typeface.GetTableTags();
            if (availableTags == null || availableTags.Length == 0)
            {
                throw new InvalidOperationException("TrueType prototype could not enumerate font tables for " + typeface.FamilyName + ".");
            }

            for (var i = 0; i < requiredTags.Length; i++)
            {
                var required = requiredTags[i];
                if (!ContainsTag(availableTags, required))
                {
                    throw new InvalidOperationException("TrueType prototype missing required font table " + TagName(required) + " for " + typeface.FamilyName + ".");
                }

                var tableData = typeface.GetTableData(required);
                if (tableData == null || tableData.Length == 0)
                {
                    throw new InvalidOperationException("TrueType prototype could not read font table " + TagName(required) + " for " + typeface.FamilyName + ".");
                }
            }
        }

        private static bool ContainsTag(uint[] tags, uint tag)
        {
            for (var i = 0; i < tags.Length; i++)
            {
                if (tags[i] == tag)
                {
                    return true;
                }
            }

            return false;
        }

        private static uint Tag(char a, char b, char c, char d)
        {
            return ((uint)a << 24) | ((uint)b << 16) | ((uint)c << 8) | d;
        }

        private static string TagName(uint tag)
        {
            var chars = new char[4];
            chars[0] = (char)((tag >> 24) & 0xFF);
            chars[1] = (char)((tag >> 16) & 0xFF);
            chars[2] = (char)((tag >> 8) & 0xFF);
            chars[3] = (char)(tag & 0xFF);
            return new string(chars);
        }

        public void Dispose()
        {
            _fallback.Dispose();
            for (var i = 0; i < _trueTypeUsages.Count; i++)
            {
                _trueTypeUsages[i].Dispose();
            }

            _trueTypeUsages.Clear();
            _probedTypefaces.Clear();
        }
    }
}
