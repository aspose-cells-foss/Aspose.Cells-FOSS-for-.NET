using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// A drawing shape reduced to what the renderer needs, parsed from the preserved shape XML
    /// (xdr:sp / xdr:cxnSp): geometry, fill, line (width + arrowheads), and text.
    /// </summary>
    internal sealed class ParsedShape
    {
        public string Geometry = "rect";
        public bool FlipH;
        public bool FlipV;
        public double RotationDeg;
        public double[] Adjusts;   // preset geometry adjust values, as fractions (0..1)
        public XElement CustomGeometry;

        public bool HasFill;
        public SKColor FillColor = SKColors.White;
        public string FillPattern;
        public SKColor PatternForeground;
        public XElement Reflection;

        public bool HasLine = true;
        public SKColor LineColor = SKColors.Black;
        public double LineWidthPt = 0.75d;

        public string HeadEnd;   // arrowhead at the start (null/none)
        public string TailEnd;   // arrowhead at the end

        public string Text;
        public SKColor TextColor = SKColors.Black;
        public double TextSizePt = 11d;
        public bool TextBold;
        public string TextVAnchor = "ctr";  // t / ctr / b
        public string TextHAlign = "ctr";   // l / ctr / r
        public string TextVerticalType = "horz";
        public bool TextWrap;
        public double TextInsetTopPt;
        public double TextInsetLeftPt;
    }

    /// <summary>A SmartArt drawing shape with its position/size (points) relative to the diagram frame.</summary>
    internal sealed class SmartArtShape
    {
        public ParsedShape Shape;
        public double XPt;
        public double YPt;
        public double WPt;
        public double HPt;
    }

    internal static class ShapeXmlParser
    {
        private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        private static readonly XNamespace Dsp = "http://schemas.microsoft.com/office/drawing/2008/diagram";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static ParsedShape Parse(string rawXml, string geometryFallback, RenderColor colors)
        {
            if (string.IsNullOrEmpty(rawXml))
            {
                return null;
            }

            XDocument doc;
            try { doc = XDocument.Parse(rawXml); }
            catch (Exception) { return FromFallback(geometryFallback); }

            var root = doc.Root;
            if (root == null)
            {
                return FromFallback(geometryFallback);
            }

            return ParseShapeElement(root, Xdr, geometryFallback, colors);
        }

        /// <summary>
        /// Parses a SmartArt diagram drawing part (dsp:drawing) into its laid-out shapes. Each shape
        /// carries an absolute position/size (points) relative to the diagram frame.
        /// </summary>
        public static List<SmartArtShape> ParseSmartArtDrawing(string rawXml, RenderColor colors)
        {
            var result = new List<SmartArtShape>();
            if (string.IsNullOrEmpty(rawXml))
            {
                return result;
            }

            XDocument doc;
            try { doc = XDocument.Parse(rawXml); }
            catch (Exception) { return result; }

            foreach (var sp in doc.Descendants(Dsp + "sp"))
            {
                var spPr = sp.Element(Dsp + "spPr");
                var xfrm = spPr != null ? spPr.Element(A + "xfrm") : null;
                if (xfrm == null)
                {
                    continue;
                }

                var off = xfrm.Element(A + "off");
                var ext = xfrm.Element(A + "ext");
                double ox, oy, cx, cy;
                if (off == null || ext == null
                    || !double.TryParse(AttrOf(off, "x"), NumberStyles.Float, CultureInfo.InvariantCulture, out ox)
                    || !double.TryParse(AttrOf(off, "y"), NumberStyles.Float, CultureInfo.InvariantCulture, out oy)
                    || !double.TryParse(AttrOf(ext, "cx"), NumberStyles.Float, CultureInfo.InvariantCulture, out cx)
                    || !double.TryParse(AttrOf(ext, "cy"), NumberStyles.Float, CultureInfo.InvariantCulture, out cy)
                    || cx <= 0d || cy <= 0d)
                {
                    continue;
                }

                var shape = ParseShapeElement(sp, Dsp, "rect", colors);
                result.Add(new SmartArtShape
                {
                    Shape = shape,
                    XPt = ox / 12700d,
                    YPt = oy / 12700d,
                    WPt = cx / 12700d,
                    HPt = cy / 12700d,
                });
            }

            return result;
        }

        private static ParsedShape ParseShapeElement(XElement root, XNamespace shapeNs, string geometryFallback, RenderColor colors)
        {
            var spPr = root.Element(shapeNs + "spPr");
            var style = root.Element(shapeNs + "style");
            var txBody = root.Element(shapeNs + "txBody");
            var shape = new ParsedShape();

            var xfrm = spPr != null ? spPr.Element(A + "xfrm") : null;
            if (xfrm != null)
            {
                shape.FlipH = string.Equals(AttrOf(xfrm, "flipH"), "1", StringComparison.Ordinal);
                shape.FlipV = string.Equals(AttrOf(xfrm, "flipV"), "1", StringComparison.Ordinal);
                double rot;
                if (double.TryParse(AttrOf(xfrm, "rot"), NumberStyles.Float, CultureInfo.InvariantCulture, out rot))
                {
                    shape.RotationDeg = rot / 60000d; // DrawingML angles are 60000ths of a degree.
                }
            }

            var prstGeom = spPr != null ? spPr.Element(A + "prstGeom") : null;
            var prst = prstGeom != null ? AttrOf(prstGeom, "prst") : null;
            shape.Geometry = !string.IsNullOrEmpty(prst) ? prst : (geometryFallback ?? "rect");
            ParseAdjusts(prstGeom, shape);
            shape.CustomGeometry = spPr != null ? spPr.Element(A + "custGeom") : null;

            ParseFill(spPr, style, colors, shape);
            ParseLine(spPr, style, colors, shape);
            ParseText(txBody, style, colors, shape);
            var effects = spPr != null ? spPr.Element(A + "effectLst") : null;
            if (effects == null && style != null && colors != null)
            {
                int effectIndex;
                if (int.TryParse(AttrOf(style.Element(A + "effectRef"), "idx"), out effectIndex))
                    effects = colors.ResolveThemeEffects(effectIndex);
            }
            shape.Reflection = effects != null ? effects.Element(A + "reflection") : null;
            return shape;
        }

        private static void ParseAdjusts(XElement prstGeom, ParsedShape shape)
        {
            var avLst = prstGeom != null ? prstGeom.Element(A + "avLst") : null;
            if (avLst == null)
            {
                return;
            }

            var values = new List<double>();
            foreach (var gd in avLst.Elements(A + "gd"))
            {
                var fmla = AttrOf(gd, "fmla");
                double v;
                if (!string.IsNullOrEmpty(fmla) && fmla.StartsWith("val ", StringComparison.Ordinal)
                    && double.TryParse(fmla.Substring(4), NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                {
                    values.Add(v / 100000d); // adjust values are in 1000ths of a percent
                }
            }
            shape.Adjusts = values.ToArray();
        }

        private static ParsedShape FromFallback(string geometry)
        {
            return new ParsedShape { Geometry = geometry ?? "rect", HasFill = false };
        }

        private static void ParseFill(XElement spPr, XElement style, RenderColor colors, ParsedShape shape)
        {
            if (spPr != null)
            {
                if (spPr.Element(A + "noFill") != null)
                {
                    shape.HasFill = false;
                    return;
                }

                var solid = spPr.Element(A + "solidFill");
                SKColor c;
                var pattern = spPr.Element(A + "pattFill");
                if (pattern != null && AttrOf(pattern, "prst") == "pct5")
                {
                    shape.HasFill = true;
                    shape.FillPattern = "pct5";
                    shape.FillColor = TryColor(FillChild(pattern.Element(A + "bgClr")), colors, out c) ? c : SKColors.White;
                    shape.PatternForeground = TryColor(FillChild(pattern.Element(A + "fgClr")), colors, out c) ? c : SKColors.Black;
                    return;
                }
                if (solid != null && TryColor(FillChild(solid), colors, out c))
                {
                    shape.HasFill = true;
                    shape.FillColor = c;
                    return;
                }
            }

            // Fall back to the style's fillRef (idx 0 means "no fill").
            var fillRef = style != null ? style.Element(A + "fillRef") : null;
            if (fillRef != null && !string.Equals(AttrOf(fillRef, "idx"), "0", StringComparison.Ordinal))
            {
                SKColor c;
                if (TryColor(FillChild2(fillRef), colors, out c))
                {
                    shape.HasFill = true;
                    shape.FillColor = c;
                }
            }
        }

        private static void ParseLine(XElement spPr, XElement style, RenderColor colors, ParsedShape shape)
        {
            var ln = spPr != null ? spPr.Element(A + "ln") : null;

            if (ln != null && ln.Element(A + "noFill") != null)
            {
                shape.HasLine = false;
                return;
            }

            // Width
            double emu;
            if (ln != null && double.TryParse(AttrOf(ln, "w"), NumberStyles.Float, CultureInfo.InvariantCulture, out emu) && emu > 0d)
            {
                shape.LineWidthPt = emu / 12700d;
            }
            else
            {
                var lnRefIdx = style != null && style.Element(A + "lnRef") != null ? AttrOf(style.Element(A + "lnRef"), "idx") : null;
                int idx;
                shape.LineWidthPt = int.TryParse(lnRefIdx, out idx) && colors != null
                    ? colors.ResolveThemeLineWidth(idx, 0.75d) : 0.75d;
            }

            // Color: explicit ln solidFill wins, then the style's lnRef.
            SKColor c;
            var lnFill = ln != null ? ln.Element(A + "solidFill") : null;
            if (lnFill != null && TryColor(FillChild(lnFill), colors, out c))
            {
                shape.HasLine = true;
                shape.LineColor = c;
            }
            else
            {
                var lnRef = style != null ? style.Element(A + "lnRef") : null;
                if (lnRef != null && TryColor(FillChild2(lnRef), colors, out c))
                {
                    shape.HasLine = true;
                    shape.LineColor = c;
                }
            }

            // Arrowheads
            if (ln != null)
            {
                var head = ln.Element(A + "headEnd");
                var tail = ln.Element(A + "tailEnd");
                shape.HeadEnd = head != null ? (AttrOf(head, "type") ?? "none") : null;
                shape.TailEnd = tail != null ? (AttrOf(tail, "type") ?? "none") : null;
            }
        }

        private static void ParseText(XElement txBody, XElement style, RenderColor colors, ParsedShape shape)
        {
            if (txBody == null)
            {
                return;
            }

            var sb = new StringBuilder();
            foreach (var t in txBody.Descendants(A + "t"))
            {
                sb.Append(t.Value);
            }
            shape.Text = sb.ToString();

            // Body properties: vertical anchor and text insets.
            var bodyPr = txBody.Element(A + "bodyPr");
            if (bodyPr != null)
            {
                var anchor = AttrOf(bodyPr, "anchor");
                if (!string.IsNullOrEmpty(anchor)) shape.TextVAnchor = anchor;
                var verticalType = AttrOf(bodyPr, "vert");
                if (!string.IsNullOrEmpty(verticalType)) shape.TextVerticalType = verticalType;
                shape.TextWrap = string.Equals(AttrOf(bodyPr, "wrap"), "square", StringComparison.Ordinal);
                double ins;
                if (double.TryParse(AttrOf(bodyPr, "tIns"), NumberStyles.Float, CultureInfo.InvariantCulture, out ins)) shape.TextInsetTopPt = ins / 12700d;
                else shape.TextInsetTopPt = 3.6d;
                if (double.TryParse(AttrOf(bodyPr, "lIns"), NumberStyles.Float, CultureInfo.InvariantCulture, out ins)) shape.TextInsetLeftPt = ins / 12700d;
                else shape.TextInsetLeftPt = 7.2d;
            }

            var pPr = txBody.Descendants(A + "pPr").FirstOrDefaultSafe();
            var algn = pPr != null ? AttrOf(pPr, "algn") : null;
            if (!string.IsNullOrEmpty(algn)) shape.TextHAlign = algn;
            else shape.TextHAlign = "l";

            // Pick up size/bold/color from the first run properties available.
            var rPr = txBody.Descendants(A + "rPr").FirstOrDefaultSafe() ?? txBody.Descendants(A + "endParaRPr").FirstOrDefaultSafe();
            if (rPr != null)
            {
                double sz;
                if (double.TryParse(AttrOf(rPr, "sz"), NumberStyles.Float, CultureInfo.InvariantCulture, out sz) && sz > 0d)
                {
                    shape.TextSizePt = sz / 100d;
                }
                shape.TextBold = string.Equals(AttrOf(rPr, "b"), "1", StringComparison.Ordinal);
                if (shape.TextBold)
                {
                    foreach (var laterRun in txBody.Descendants(A + "r"))
                    {
                        var laterText = laterRun.Element(A + "t");
                        var laterProperties = laterRun.Element(A + "rPr");
                        if (laterText != null && laterText.Value.Length > shape.Text.Length / 2
                            && laterProperties != null && !string.Equals(AttrOf(laterProperties, "b"), "1", StringComparison.Ordinal))
                        {
                            shape.TextBold = false;
                            break;
                        }
                    }
                }
                SKColor c;
                var solid = rPr.Element(A + "solidFill");
                if (solid != null && TryColor(FillChild(solid), colors, out c))
                {
                    shape.TextColor = c;
                    return;
                }
            }

            SKColor fallbackColor;
            var fontRef = style != null ? style.Element(A + "fontRef") : null;
            if (fontRef != null && TryColor(FillChild2(fontRef), colors, out fallbackColor))
            {
                shape.TextColor = fallbackColor;
            }
        }

        private static XElement FillChild(XElement solidFill)
        {
            return solidFill == null ? null : (solidFill.Element(A + "srgbClr") ?? solidFill.Element(A + "schemeClr"));
        }

        private static XElement FillChild2(XElement refEl)
        {
            // fillRef/lnRef carry a scheme or srgb color child directly.
            return refEl == null ? null : (refEl.Element(A + "srgbClr") ?? refEl.Element(A + "schemeClr"));
        }

        private static bool TryColor(XElement colorEl, RenderColor colors, out SKColor color)
        {
            color = SKColors.Black;
            if (colorEl == null)
            {
                return false;
            }

            if (colorEl.Name == A + "srgbClr")
            {
                if (TryHex(AttrOf(colorEl, "val"), out color))
                {
                    color = ApplyColorModifiers(colorEl, color);
                    return true;
                }

                return false;
            }

            if (colorEl.Name == A + "schemeClr")
            {
                var baseColor = colors.ResolveSchemeName(AttrOf(colorEl, "val"), SKColors.Gray);
                double lumMod, lumOff;
                double.TryParse(AttrOf(colorEl.Element(A + "lumMod"), "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out lumMod);
                double.TryParse(AttrOf(colorEl.Element(A + "lumOff"), "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out lumOff);
                color = (lumMod > 0d || lumOff > 0d)
                    ? RenderColor.ApplyLuma(baseColor, lumMod > 0d ? lumMod / 100000d : 1d, lumOff / 100000d)
                    : baseColor;
                color = ApplyColorModifiers(colorEl, color);
                return true;
            }

            return false;
        }

        private static SKColor ApplyColorModifiers(XElement colorEl, SKColor color)
        {
            if (colorEl == null)
            {
                return color;
            }

            var shade = AttrOf(colorEl.Element(A + "shade"), "val");
            int shadeValue;
            if (int.TryParse(shade, NumberStyles.Integer, CultureInfo.InvariantCulture, out shadeValue))
            {
                var factor = shadeValue / 100000d;
                color = new SKColor(
                    ScaleByte(color.Red, factor),
                    ScaleByte(color.Green, factor),
                    ScaleByte(color.Blue, factor),
                    color.Alpha);
            }

            var tint = AttrOf(colorEl.Element(A + "tint"), "val");
            int tintValue;
            if (int.TryParse(tint, NumberStyles.Integer, CultureInfo.InvariantCulture, out tintValue))
            {
                var amount = tintValue / 100000d;
                color = new SKColor(
                    TintByte(color.Red, amount),
                    TintByte(color.Green, amount),
                    TintByte(color.Blue, amount),
                    color.Alpha);
            }

            int alpha;
            if (int.TryParse(AttrOf(colorEl.Element(A + "alpha"), "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out alpha))
            {
                color = color.WithAlpha((byte)Math.Round(Math.Max(0, Math.Min(100000, alpha)) * 255d / 100000d));
            }

            return color;
        }

        private static byte ScaleByte(byte value, double factor)
        {
            // DrawingML shade operates on linear light, not encoded sRGB bytes.
            var channel = value / 255d;
            var linear = channel <= 0.04045d ? channel / 12.92d : Math.Pow((channel + 0.055d) / 1.055d, 2.4d);
            linear = Math.Max(0d, Math.Min(1d, linear * factor));
            var encoded = linear <= 0.0031308d ? linear * 12.92d : 1.055d * Math.Pow(linear, 1d / 2.4d) - 0.055d;
            return (byte)Math.Round(encoded * 255d);
        }

        private static byte TintByte(byte value, double amount)
        {
            var tinted = value + (255d - value) * amount;
            if (tinted < 0d) tinted = 0d;
            if (tinted > 255d) tinted = 255d;
            return (byte)Math.Round(tinted);
        }

        private static bool TryHex(string hex, out SKColor color)
        {
            color = SKColors.Black;
            if (string.IsNullOrEmpty(hex) || hex.Length < 6)
            {
                return false;
            }

            var offset = hex.Length == 8 ? 2 : 0;
            int r, g, b;
            if (int.TryParse(hex.Substring(offset, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r)
                && int.TryParse(hex.Substring(offset + 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g)
                && int.TryParse(hex.Substring(offset + 4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b))
            {
                color = new SKColor((byte)r, (byte)g, (byte)b);
                return true;
            }

            return false;
        }

        private static string AttrOf(XElement e, string name)
        {
            if (e == null) return null;
            var a = e.Attribute(name);
            return a != null ? a.Value : null;
        }
    }
}
