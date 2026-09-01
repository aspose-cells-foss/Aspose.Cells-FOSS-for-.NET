using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// A single typeface-homogeneous slice of a text string, with its measured advance width in
    /// points. Produced by <see cref="CellFontContext.SplitRuns"/> and consumed both when measuring
    /// (layout) and drawing (rendering) so the two stay in lock-step.
    /// </summary>
    internal sealed class GlyphRun
    {
        public SKTypeface Typeface;
        public string Text;
        public float WidthPt;
    }

    /// <summary>
    /// Wraps a resolved font configuration. All measurement here is in points (SKFont size equals the
    /// point size, which matches the renderer's 72-DPI PDF canvas), so widths are directly comparable
    /// to the point-based column widths from the layout layer.
    /// </summary>
    internal sealed class CellFontContext : IDisposable
    {
        private readonly FontRegistry _registry;
        private readonly SKTypeface _primary;
        private readonly float _sizePt;
        private readonly Dictionary<SKTypeface, SKFont> _fonts = new Dictionary<SKTypeface, SKFont>();

        public CellFontContext(FontRegistry registry, FontValue font)
        {
            _registry = registry;
            _primary = registry.Resolve(font);
            _sizePt = (float)(font != null ? font.Size : 11d);
            if (_sizePt <= 0f)
            {
                _sizePt = 11f;
            }
        }

        public float SizePt { get { return _sizePt; } }

        /// <summary>Approximate line height in points (Excel adds ~20% leading over the em size).</summary>
        public float LineHeightPt { get { return _sizePt * 1.2f; } }

        /// <summary>
        /// Glyph height (ascent + descent, no leading) in points, matching the vertical extent the
        /// renderer draws. Used to size a row to rotated text.
        /// </summary>
        public float TextHeightPt
        {
            get
            {
                var metrics = FontFor(_primary).Metrics;
                return -metrics.Ascent + metrics.Descent;
            }
        }

        /// <summary>
        /// Tight ink bounds of <paramref name="text"/> (the actual drawn extent, excluding the
        /// descender space the font metrics reserve). Used to size a row to rotated text the way Excel
        /// does - to the glyphs present, not the font's nominal line box.
        /// </summary>
        public void MeasureInkBounds(string text, out float width, out float height)
        {
            width = 0f;
            height = 0f;
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var bounds = new SKRect();
            using (var paint = new SKPaint(FontFor(_primary)))
            {
                paint.MeasureText(text, ref bounds);
            }

            width = bounds.Width;
            height = bounds.Height;
            if (height <= 0f)
            {
                // Fall back to the metric height when a face reports no ink (e.g. all-whitespace).
                height = TextHeightPt;
            }
        }

        private SKFont FontFor(SKTypeface typeface)
        {
            SKFont font;
            if (_fonts.TryGetValue(typeface, out font))
            {
                return font;
            }

            font = new SKFont(typeface, _sizePt);
            font.Subpixel = true;
            _fonts[typeface] = font;
            return font;
        }

        /// <summary>
        /// Splits <paramref name="text"/> into runs, each covered by a single typeface (primary where
        /// possible, glyph-coverage fallback otherwise).
        /// </summary>
        public List<GlyphRun> SplitRuns(string text)
        {
            var runs = new List<GlyphRun>();
            if (string.IsNullOrEmpty(text))
            {
                return runs;
            }

            var builder = new System.Text.StringBuilder();
            SKTypeface currentFace = null;

            var index = 0;
            while (index < text.Length)
            {
                int codepoint;
                int consumed = ReadCodepoint(text, index, out codepoint);
                var face = _registry.ResolveForCodepoint(_primary, codepoint);

                if (currentFace == null)
                {
                    currentFace = face;
                }
                else if (!ReferenceEquals(face, currentFace))
                {
                    runs.Add(MakeRun(currentFace, builder.ToString()));
                    builder.Length = 0;
                    currentFace = face;
                }

                builder.Append(text, index, consumed);
                index += consumed;
            }

            if (builder.Length > 0)
            {
                runs.Add(MakeRun(currentFace, builder.ToString()));
            }

            return runs;
        }

        private GlyphRun MakeRun(SKTypeface typeface, string text)
        {
            var font = FontFor(typeface);
            using (var paint = new SKPaint(font))
            {
                return new GlyphRun { Typeface = typeface, Text = text, WidthPt = paint.MeasureText(text) };
            }
        }

        public float Measure(string text)
        {
            var runs = SplitRuns(text);
            float total = 0f;
            for (var i = 0; i < runs.Count; i++)
            {
                total += runs[i].WidthPt;
            }

            return total;
        }

        /// <summary>
        /// Greedy line breaker: wraps on whitespace where possible, hard-breaking any single token
        /// that is wider than <paramref name="maxWidthPt"/>. Explicit newlines are always honored.
        /// </summary>
        public List<string> WrapLines(string text, float maxWidthPt)
        {
            var lines = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                lines.Add(string.Empty);
                return lines;
            }

            var paragraphs = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (var p = 0; p < paragraphs.Length; p++)
            {
                WrapParagraph(paragraphs[p], maxWidthPt, lines);
            }

            return lines;
        }

        private void WrapParagraph(string paragraph, float maxWidthPt, List<string> lines)
        {
            if (maxWidthPt <= 0f || paragraph.Length == 0)
            {
                lines.Add(paragraph);
                return;
            }

            var current = new System.Text.StringBuilder();
            var tokens = Tokenize(paragraph);
            for (var t = 0; t < tokens.Count; t++)
            {
                var token = tokens[t];
                var candidate = current.Length == 0 ? token : current.ToString() + token;
                if (Measure(candidate) <= maxWidthPt || current.Length == 0 && Measure(token) <= maxWidthPt)
                {
                    current.Append(token);
                    continue;
                }

                if (current.Length > 0)
                {
                    lines.Add(current.ToString().TrimEnd());
                    current.Length = 0;
                }

                // Token alone still overflows: hard-break it character by character.
                if (Measure(token) > maxWidthPt)
                {
                    HardBreak(token, maxWidthPt, lines, current);
                }
                else
                {
                    current.Append(token);
                }
            }

            lines.Add(current.ToString().TrimEnd());
        }

        private void HardBreak(string token, float maxWidthPt, List<string> lines, System.Text.StringBuilder current)
        {
            foreach (var ch in token)
            {
                var candidate = current.ToString() + ch;
                if (current.Length > 0 && Measure(candidate) > maxWidthPt)
                {
                    lines.Add(current.ToString());
                    current.Length = 0;
                }

                current.Append(ch);
            }
        }

        private static List<string> Tokenize(string paragraph)
        {
            // Keep trailing whitespace attached to each word so widths stay accurate.
            var tokens = new List<string>();
            var builder = new System.Text.StringBuilder();
            for (var i = 0; i < paragraph.Length; i++)
            {
                var ch = paragraph[i];
                builder.Append(ch);
                if (ch == ' ' || ch == '\t')
                {
                    tokens.Add(builder.ToString());
                    builder.Length = 0;
                }
            }

            if (builder.Length > 0)
            {
                tokens.Add(builder.ToString());
            }

            return tokens;
        }

        private static int ReadCodepoint(string text, int index, out int codepoint)
        {
            var ch = text[index];
            if (char.IsHighSurrogate(ch) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]))
            {
                codepoint = char.ConvertToUtf32(ch, text[index + 1]);
                return 2;
            }

            codepoint = ch;
            return 1;
        }

        public void Dispose()
        {
            foreach (var font in _fonts.Values)
            {
                font.Dispose();
            }

            _fonts.Clear();
        }
    }
}
