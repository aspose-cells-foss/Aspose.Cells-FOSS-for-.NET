using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Resolves a <see cref="ColorValue"/> (which may be a direct ARGB value, a theme reference
    /// with an optional tint, or a legacy indexed color) into a concrete <see cref="SKColor"/>.
    /// </summary>
    internal sealed class RenderColor
    {
        // Office default theme palette, stored in clrScheme document order:
        // dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink.
        private static readonly SKColor[] DefaultScheme =
        {
            new SKColor(0x00, 0x00, 0x00), // dk1 (text1)
            new SKColor(0xFF, 0xFF, 0xFF), // lt1 (bg1)
            new SKColor(0x44, 0x54, 0x6A), // dk2 (text2)
            new SKColor(0xE7, 0xE6, 0xE6), // lt2 (bg2)
            new SKColor(0x44, 0x72, 0xC4), // accent1
            new SKColor(0xED, 0x7D, 0x31), // accent2
            new SKColor(0xA5, 0xA5, 0xA5), // accent3
            new SKColor(0xFF, 0xC0, 0x00), // accent4
            new SKColor(0x5B, 0x9B, 0xD5), // accent5
            new SKColor(0x70, 0xAD, 0x47), // accent6
            new SKColor(0x05, 0x63, 0xC1), // hlink
            new SKColor(0x95, 0x4F, 0x72), // folHlink
        };

        // Classic 56-entry indexed color palette (indices 8..63). Indices 0..7 mirror 8..15.
        private static readonly SKColor[] IndexedPalette = BuildIndexedPalette();

        private readonly SKColor[] _scheme;

        private RenderColor(SKColor[] scheme)
        {
            _scheme = scheme;
        }

        /// <summary>
        /// Builds a resolver from the workbook's raw theme XML (may be null, in which case the
        /// Office default palette is used).
        /// </summary>
        public static RenderColor FromWorkbook(WorkbookModel model)
        {
            var scheme = ParseThemeScheme(model != null ? model.RawThemeXml : null);
            return new RenderColor(scheme ?? DefaultScheme);
        }

        /// <summary>
        /// Resolves the color; returns <paramref name="fallback"/> when the value is empty/unset.
        /// </summary>
        public SKColor Resolve(ColorValue color, SKColor fallback)
        {
            if (color.ThemeIndex.HasValue)
            {
                var baseColor = ResolveTheme(color.ThemeIndex.Value);
                return ApplyTint(baseColor, color.Tint ?? 0d);
            }

            if (color.Indexed.HasValue)
            {
                return ResolveIndexed(color.Indexed.Value, fallback);
            }

            if (color.A == 0 && color.R == 0 && color.G == 0 && color.B == 0)
            {
                // A fully-zero, non-theme, non-indexed value is the "unset" sentinel.
                return fallback;
            }

            return new SKColor(color.R, color.G, color.B, color.A);
        }

        /// <summary>
        /// Resolves a DrawingML scheme color name (e.g. "accent1", "tx1", "bg1") to its theme color.
        /// Used by chart rendering, where series and text reference the theme by name.
        /// </summary>
        public SKColor ResolveSchemeName(string name, SKColor fallback)
        {
            int index;
            switch (name)
            {
                case "dk1": case "tx1": index = 0; break;
                case "lt1": case "bg1": index = 1; break;
                case "dk2": case "tx2": index = 2; break;
                case "lt2": case "bg2": index = 3; break;
                case "accent1": index = 4; break;
                case "accent2": index = 5; break;
                case "accent3": index = 6; break;
                case "accent4": index = 7; break;
                case "accent5": index = 8; break;
                case "accent6": index = 9; break;
                case "hlink": index = 10; break;
                case "folHlink": index = 11; break;
                default: return fallback;
            }

            return index >= 0 && index < _scheme.Length ? _scheme[index] : fallback;
        }

        /// <summary>Applies a DrawingML luminance modulation/offset (lumMod/lumOff) to a color.</summary>
        public static SKColor ApplyLuma(SKColor color, double lumMod, double lumOff)
        {
            return new SKColor(
                LumaChannel(color.Red, lumMod, lumOff),
                LumaChannel(color.Green, lumMod, lumOff),
                LumaChannel(color.Blue, lumMod, lumOff),
                color.Alpha);
        }

        private static byte LumaChannel(byte channel, double lumMod, double lumOff)
        {
            var value = channel * lumMod + 255d * lumOff;
            if (value < 0d) value = 0d;
            if (value > 255d) value = 255d;
            return (byte)Math.Round(value);
        }

        private SKColor ResolveTheme(int themeIndex)
        {
            // The style theme index order swaps the first two dark/light pairs relative to the
            // clrScheme document order: 0=lt1, 1=dk1, 2=lt2, 3=dk2.
            int schemeIndex;
            switch (themeIndex)
            {
                case 0: schemeIndex = 1; break; // lt1
                case 1: schemeIndex = 0; break; // dk1
                case 2: schemeIndex = 3; break; // lt2
                case 3: schemeIndex = 2; break; // dk2
                default: schemeIndex = themeIndex; break;
            }

            if (schemeIndex >= 0 && schemeIndex < _scheme.Length)
            {
                return _scheme[schemeIndex];
            }

            return SKColors.Black;
        }

        private static SKColor ResolveIndexed(int index, SKColor fallback)
        {
            if (index == 64 || index == 65)
            {
                // 64 = system foreground, 65 = system background; caller supplies the right default.
                return fallback;
            }

            if (index >= 0 && index < IndexedPalette.Length)
            {
                return IndexedPalette[index];
            }

            return fallback;
        }

        /// <summary>
        /// Applies Excel's tint/shade to a base color. Negative values darken, positive lighten.
        /// This is the widely-used per-channel RGB approximation of Excel's HLS-based tinting.
        /// </summary>
        private static SKColor ApplyTint(SKColor color, double tint)
        {
            if (tint == 0d)
            {
                return color;
            }

            return new SKColor(TintChannel(color.Red, tint), TintChannel(color.Green, tint), TintChannel(color.Blue, tint), color.Alpha);
        }

        private static byte TintChannel(byte channel, double tint)
        {
            double value;
            if (tint < 0d)
            {
                value = channel * (1d + tint);
            }
            else
            {
                value = channel * (1d - tint) + 255d * tint;
            }

            if (value < 0d) value = 0d;
            if (value > 255d) value = 255d;
            return (byte)Math.Round(value);
        }

        private static SKColor[] ParseThemeScheme(string rawThemeXml)
        {
            if (string.IsNullOrWhiteSpace(rawThemeXml))
            {
                return null;
            }

            try
            {
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                var doc = XDocument.Parse(rawThemeXml);
                var clrScheme = doc.Descendants(a + "clrScheme").FirstOrDefaultSafe();
                if (clrScheme == null)
                {
                    return null;
                }

                var order = new[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
                var result = new SKColor[order.Length];
                for (var i = 0; i < order.Length; i++)
                {
                    var element = clrScheme.Element(a + order[i]);
                    SKColor parsed;
                    if (element != null && TryParseSchemeColor(element, a, out parsed))
                    {
                        result[i] = parsed;
                    }
                    else
                    {
                        result[i] = DefaultScheme[i];
                    }
                }

                return result;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryParseSchemeColor(XElement colorSlot, XNamespace a, out SKColor color)
        {
            color = SKColors.Black;

            var srgb = colorSlot.Element(a + "srgbClr");
            if (srgb != null)
            {
                var val = (string)srgb.Attribute("val");
                return TryParseHex(val, out color);
            }

            var sys = colorSlot.Element(a + "sysClr");
            if (sys != null)
            {
                // sysClr carries a lastClr attribute with the resolved RGB.
                var last = (string)sys.Attribute("lastClr");
                if (TryParseHex(last, out color))
                {
                    return true;
                }

                var name = (string)sys.Attribute("val");
                if (string.Equals(name, "window", StringComparison.OrdinalIgnoreCase))
                {
                    color = SKColors.White;
                    return true;
                }

                if (string.Equals(name, "windowText", StringComparison.OrdinalIgnoreCase))
                {
                    color = SKColors.Black;
                    return true;
                }
            }

            return false;
        }

        private static bool TryParseHex(string hex, out SKColor color)
        {
            color = SKColors.Black;
            if (string.IsNullOrEmpty(hex) || hex.Length < 6)
            {
                return false;
            }

            int offset = hex.Length == 8 ? 2 : 0; // tolerate leading alpha
            byte r, g, b;
            if (TryHexByte(hex, offset, out r) && TryHexByte(hex, offset + 2, out g) && TryHexByte(hex, offset + 4, out b))
            {
                color = new SKColor(r, g, b);
                return true;
            }

            return false;
        }

        private static bool TryHexByte(string s, int index, out byte value)
        {
            value = 0;
            if (index + 2 > s.Length)
            {
                return false;
            }

            int result;
            if (int.TryParse(s.Substring(index, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out result))
            {
                value = (byte)result;
                return true;
            }

            return false;
        }

        private static SKColor[] BuildIndexedPalette()
        {
            // Excel's legacy default color index table. Values are 0xRRGGBB.
            uint[] rgb =
            {
                0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
                0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
                0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
                0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
                0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
                0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
                0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
                0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
            };

            var palette = new SKColor[rgb.Length];
            for (var i = 0; i < rgb.Length; i++)
            {
                palette[i] = new SKColor((byte)((rgb[i] >> 16) & 0xFF), (byte)((rgb[i] >> 8) & 0xFF), (byte)(rgb[i] & 0xFF));
            }

            return palette;
        }
    }

    internal static class XElementExtensions
    {
        public static XElement FirstOrDefaultSafe(this IEnumerable<XElement> source)
        {
            foreach (var element in source)
            {
                return element;
            }

            return null;
        }
    }
}
