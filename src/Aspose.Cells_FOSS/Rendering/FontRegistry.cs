using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Caches <see cref="SKTypeface"/> instances and provides a glyph-coverage fallback chain so
    /// text (including CJK) renders without depending on a specific set of installed fonts.
    /// </summary>
    internal sealed class FontRegistry : IDisposable
    {
        private readonly Dictionary<string, SKTypeface> _cache = new Dictionary<string, SKTypeface>(StringComparer.Ordinal);
        private readonly List<SKTypeface> _fallbacks = new List<SKTypeface>();
        private bool _fallbacksInitialized;

        /// <summary>
        /// Per-font Maximum Digit Width (pixels at 11pt/96 DPI) for fonts where Excel/GDI measures the
        /// digit advance differently from SkiaSharp's outline metrics - typically CJK fonts that use
        /// embedded bitmap glyphs at UI sizes. Keyed by both the OOXML family name and the resolved
        /// SkiaSharp family name. Empirically calibrated against Excel's own PDF exports.
        /// </summary>
        private static readonly Dictionary<string, double> ExcelDigitWidthPxAt11 = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
        {
            { "SimSun", 7.75d },
            { "宋体", 7.75d },   // 宋体
            { "NSimSun", 7.75d },
            { "新宋体", 7.75d }, // 新宋体
        };

        /// <summary>
        /// Optional user-supplied fallback font family (from PdfSaveOptions.DefaultFont). When set,
        /// it is tried before the built-in fallback chain during glyph-coverage resolution.
        /// </summary>
        public string DefaultFontName { get; set; }

        /// <summary>
        /// Resolves the primary typeface for a font, honoring name/weight/slant. Falls back to the
        /// default system typeface when the requested family is unavailable.
        /// </summary>
        public SKTypeface Resolve(FontValue font)
        {
            var name = font != null && !string.IsNullOrEmpty(font.Name) ? font.Name : "Calibri";
            var weight = (font != null && font.Bold) ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
            var slant = (font != null && font.Italic) ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
            var key = name + "|" + (int)weight + "|" + (int)slant;

            SKTypeface cached;
            if (_cache.TryGetValue(key, out cached))
            {
                return cached;
            }

            var typeface = SKTypeface.FromFamilyName(name, weight, SKFontStyleWidth.Normal, slant)
                ?? SKTypeface.Default;
            _cache[key] = typeface;
            return typeface;
        }

        /// <summary>
        /// Returns a typeface that contains a glyph for <paramref name="codepoint"/>, preferring the
        /// primary typeface and walking the platform fallback set otherwise.
        /// </summary>
        public SKTypeface ResolveForCodepoint(SKTypeface primary, int codepoint)
        {
            if (primary != null && primary.ContainsGlyph(codepoint))
            {
                return primary;
            }

            EnsureFallbacks();
            for (var i = 0; i < _fallbacks.Count; i++)
            {
                if (_fallbacks[i] != null && _fallbacks[i].ContainsGlyph(codepoint))
                {
                    return _fallbacks[i];
                }
            }

            // Ask Skia's font manager for a family that covers this specific character.
            var matched = SKFontManager.Default.MatchCharacter(codepoint);
            if (matched != null)
            {
                _fallbacks.Add(matched);
                return matched;
            }

            return primary ?? SKTypeface.Default;
        }

        private void EnsureFallbacks()
        {
            if (_fallbacksInitialized)
            {
                return;
            }

            _fallbacksInitialized = true;

            // A user-supplied default font takes precedence over the built-in candidates.
            if (!string.IsNullOrEmpty(DefaultFontName))
            {
                var preferred = SKTypeface.FromFamilyName(DefaultFontName);
                if (preferred != null)
                {
                    _fallbacks.Add(preferred);
                }
            }

            // Seed with a few widely-available CJK/Unicode families; missing ones resolve to null
            // and are skipped. MatchCharacter still covers anything these miss.
            var candidates = new[]
            {
                "Noto Sans CJK SC", "Noto Sans CJK TC", "Noto Sans CJK JP", "Noto Sans",
                "Microsoft YaHei", "SimSun", "PingFang SC", "Hiragino Sans", "Arial Unicode MS",
            };

            foreach (var candidate in candidates)
            {
                var typeface = SKTypeface.FromFamilyName(candidate);
                if (typeface != null && !string.Equals(typeface.FamilyName, SKTypeface.Default.FamilyName, StringComparison.Ordinal))
                {
                    _fallbacks.Add(typeface);
                }
            }
        }

        /// <summary>
        /// Measures the Maximum Digit Width (width of the glyph '0') for the workbook's normal font,
        /// in pixels at 96 DPI. Used to convert Excel character-based column widths to device units.
        /// </summary>
        public double MeasureMaxDigitWidth(FontValue normalFont)
        {
            var typeface = Resolve(normalFont);
            var sizePt = normalFont != null ? normalFont.Size : 11d;

            // A calibrated override wins over measurement for fonts SkiaSharp mis-measures vs Excel.
            double overridePxAt11;
            if (TryGetDigitWidthOverride(normalFont, typeface, out overridePxAt11))
            {
                return overridePxAt11 * (sizePt / 11d);
            }

            var sizePx = sizePt * RenderUnits.PointsToPixels;
            using (var font = new SKFont(typeface, (float)sizePx))
            {
                // Excel derives column widths from the font's true (unhinted) digit advance. The
                // default hinted measurement snaps to whole pixels (e.g. Calibri 11 -> 7.0 instead
                // of 7.43), which yields visibly narrower columns than Excel; use linear metrics.
                font.Subpixel = true;
                font.Hinting = SKFontHinting.None;
                font.LinearMetrics = true;
                using (var paint = new SKPaint(font))
                {
                    var width = paint.MeasureText("0");
                    if (width <= 0f)
                    {
                        // Calibri 11 default fallback.
                        return 7.5d;
                    }

                    // Excel's effective digit metric sits on a half-pixel grid at or just above the
                    // raw typographic advance (Calibri 11: 7.43 -> 7.5; DengXian 11: 7.73 -> 8.0).
                    // Rounding up to the next half pixel matches Excel's column widths closely across
                    // both Latin and CJK default fonts.
                    return Math.Ceiling(width * 2d) / 2d;
                }
            }
        }

        private static bool TryGetDigitWidthOverride(FontValue font, SKTypeface typeface, out double pxAt11)
        {
            if (font != null && !string.IsNullOrEmpty(font.Name))
            {
                if (string.Equals(font.Name, "DengXian", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(font.Name, "等线", StringComparison.Ordinal))
                {
                    pxAt11 = 7.75d;
                    return true;
                }
            }

            if (typeface != null && !string.IsNullOrEmpty(typeface.FamilyName))
            {
                if (string.Equals(typeface.FamilyName, "DengXian", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(typeface.FamilyName, "等线", StringComparison.Ordinal))
                {
                    pxAt11 = 7.75d;
                    return true;
                }
            }

            if (font != null && !string.IsNullOrEmpty(font.Name) && ExcelDigitWidthPxAt11.TryGetValue(font.Name, out pxAt11))
            {
                return true;
            }

            if (typeface != null && !string.IsNullOrEmpty(typeface.FamilyName) && ExcelDigitWidthPxAt11.TryGetValue(typeface.FamilyName, out pxAt11))
            {
                return true;
            }

            pxAt11 = 0d;
            return false;
        }

        /// <summary>
        /// The natural row height (in points) Excel derives from a font: the font's full line spacing
        /// (ascent + descent + leading) plus ~1pt of cell padding. Including the font's leading is
        /// what matches Excel across Latin fonts with no leading (Calibri 11 -> ~14.4, DengXian 11 ->
        /// ~12.48) and CJK fonts that carry line-gap leading (SimSun 11 -> ~13.5).
        /// </summary>
        public double MeasureDefaultRowHeightPt(FontValue normalFont)
        {
            var typeface = Resolve(normalFont);
            var sizePt = (float)(normalFont != null ? normalFont.Size : 11d);
            using (var font = new SKFont(typeface, sizePt))
            {
                var metrics = font.Metrics;
                var lineHeight = -metrics.Ascent + metrics.Descent + metrics.Leading; // full line spacing
                if (lineHeight <= 0f)
                {
                    return 14.4d;
                }

                return lineHeight + 1.0d;
            }
        }

        public void Dispose()
        {
            foreach (var typeface in _cache.Values)
            {
                if (typeface != null && !ReferenceEquals(typeface, SKTypeface.Default))
                {
                    typeface.Dispose();
                }
            }

            _cache.Clear();

            foreach (var typeface in _fallbacks)
            {
                if (typeface != null)
                {
                    typeface.Dispose();
                }
            }

            _fallbacks.Clear();
        }
    }
}
