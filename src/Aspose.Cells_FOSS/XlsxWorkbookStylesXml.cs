using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;
using static Aspose.Cells_FOSS.XlsxWorkbookStylesValueHelpers;
namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookStylesXml
    {
        internal static XDocument BuildStylesheetDocument(
            IReadOnlyList<FontValue> fonts,
            IReadOnlyList<FillValue> fills,
            IReadOnlyList<BordersValue> borders,
            CellFormatValue normalCellFormat,
            IReadOnlyList<CellFormatValue> cellFormats,
            IReadOnlyList<KeyValuePair<int, string>> customNumberFormats,
            IReadOnlyList<StyleValue> differentialFormats)
        {
            var stylesheet = new XElement(MainNs + "styleSheet");
            if (customNumberFormats.Count > 0)
            {
                var numFmtElements = new List<XElement>(customNumberFormats.Count);
                for (var index = 0; index < customNumberFormats.Count; index++)
                {
                    var pair = customNumberFormats[index];
                    numFmtElements.Add(new XElement(MainNs + "numFmt",
                        new XAttribute("numFmtId", pair.Key),
                        new XAttribute("formatCode", pair.Value)));
                }
                stylesheet.Add(new XElement(MainNs + "numFmts",
                    new XAttribute("count", customNumberFormats.Count),
                    numFmtElements));
            }
            stylesheet.Add(new XElement(MainNs + "fonts",
                new XAttribute("count", fonts.Count),
                BuildFontElements(fonts)));
            stylesheet.Add(new XElement(MainNs + "fills",
                new XAttribute("count", fills.Count),
                BuildFillElements(fills)));
            stylesheet.Add(new XElement(MainNs + "borders",
                new XAttribute("count", borders.Count),
                BuildBorderElements(borders)));
            stylesheet.Add(new XElement(MainNs + "cellStyleXfs",
                new XAttribute("count", 1),
                BuildCellStyleFormatElement(normalCellFormat)));
            stylesheet.Add(new XElement(MainNs + "cellXfs",
                new XAttribute("count", cellFormats.Count),
                BuildCellFormatElements(cellFormats)));
            stylesheet.Add(new XElement(MainNs + "cellStyles",
                new XAttribute("count", 1),
                new XElement(MainNs + "cellStyle",
                    new XAttribute("name", "Normal"),
                    new XAttribute("xfId", 0),
                    new XAttribute("builtinId", 0))));
            if (differentialFormats.Count > 0)
            {
                stylesheet.Add(new XElement(MainNs + "dxfs",
                    new XAttribute("count", differentialFormats.Count),
                    BuildDifferentialFormatElements(differentialFormats)));
            }
            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), stylesheet);
        }
        internal static List<FontValue> ReadFontValues(XElement root)
        {
            var fonts = new List<FontValue>();
            foreach (var font in root.Element(MainNs + "fonts")?.Elements(MainNs + "font") ?? Enumerable.Empty<XElement>())
            {
                fonts.Add(ReadFontValue(font));
            }
            return fonts;
        }
        internal static List<FillValue> ReadFillValues(XElement root)
        {
            var fills = new List<FillValue>();
            foreach (var fill in root.Element(MainNs + "fills")?.Elements(MainNs + "fill") ?? Enumerable.Empty<XElement>())
            {
                fills.Add(ReadFillValue(fill));
            }
            return fills;
        }
        internal static List<BordersValue> ReadBordersValues(XElement root)
        {
            var borders = new List<BordersValue>();
            foreach (var border in root.Element(MainNs + "borders")?.Elements(MainNs + "border") ?? Enumerable.Empty<XElement>())
            {
                borders.Add(ReadBordersValue(border));
            }
            return borders;
        }
        internal static List<StyleValue> ReadDifferentialStyleValues(XElement root)
        {
            var styles = new List<StyleValue>();
            foreach (var dxf in root.Element(MainNs + "dxfs")?.Elements(MainNs + "dxf") ?? Enumerable.Empty<XElement>())
            {
                styles.Add(ReadDifferentialStyleValue(dxf));
            }
            return styles;
        }
        internal static List<XElement> BuildFontElements(IReadOnlyList<FontValue> fonts)
        {
            var elements = new List<XElement>(fonts.Count);
            for (var index = 0; index < fonts.Count; index++)
            {
                elements.Add(BuildFontElement(fonts[index]));
            }
            return elements;
        }
        internal static List<XElement> BuildFillElements(IReadOnlyList<FillValue> fills)
        {
            var elements = new List<XElement>(fills.Count);
            for (var index = 0; index < fills.Count; index++)
            {
                elements.Add(BuildFillElement(fills[index]));
            }
            return elements;
        }
        internal static List<XElement> BuildBorderElements(IReadOnlyList<BordersValue> borders)
        {
            var elements = new List<XElement>(borders.Count);
            for (var index = 0; index < borders.Count; index++)
            {
                elements.Add(BuildBorderElement(borders[index]));
            }
            return elements;
        }
        internal static List<XElement> BuildCellFormatElements(IReadOnlyList<CellFormatValue> cellFormats)
        {
            var elements = new List<XElement>(cellFormats.Count);
            for (var index = 0; index < cellFormats.Count; index++)
            {
                elements.Add(BuildCellFormatElement(cellFormats[index]));
            }
            return elements;
        }
        internal static List<XElement> BuildDifferentialFormatElements(IReadOnlyList<StyleValue> differentialFormats)
        {
            var elements = new List<XElement>(differentialFormats.Count);
            for (var index = 0; index < differentialFormats.Count; index++)
            {
                elements.Add(BuildDifferentialFormatElement(differentialFormats[index]));
            }
            return elements;
        }
        internal static HorizontalAlignment ParseHorizontalAlignment(string value)
        {
            switch (value?.ToLowerInvariant())
            {
                case "left":
                    return HorizontalAlignment.Left;
                case "center":
                    return HorizontalAlignment.Center;
                case "right":
                    return HorizontalAlignment.Right;
                case "fill":
                    return HorizontalAlignment.Fill;
                case "justify":
                    return HorizontalAlignment.Justify;
                case "centercontinuous":
                    return HorizontalAlignment.CenterContinuous;
                case "distributed":
                    return HorizontalAlignment.Distributed;
                default:
                    return HorizontalAlignment.General;
            }
        }
        internal static VerticalAlignment ParseVerticalAlignment(string value)
        {
            switch (value?.ToLowerInvariant())
            {
                case "center":
                    return VerticalAlignment.Center;
                case "top":
                    return VerticalAlignment.Top;
                case "justify":
                    return VerticalAlignment.Justify;
                case "distributed":
                    return VerticalAlignment.Distributed;
                default:
                    return VerticalAlignment.Bottom;
            }
        }
        internal static bool? ParseOptionalBoolAttribute(XAttribute attribute)
        {
            if (attribute == null) return null;
            return (bool?)ParseBoolAttribute(attribute);
        }
        private static XElement BuildCellStyleFormatElement(CellFormatValue cellFormat)
        {
            return BuildCellFormatElement(cellFormat, false);
        }
        private static StyleValue ReadDifferentialStyleValue(XElement dxf)
        {
            var style = StyleValue.Default.Clone();
            var font = dxf.Element(MainNs + "font");
            if (font != null)
            {
                style.Font = ReadFontValue(font);
            }
            var fill = dxf.Element(MainNs + "fill");
            if (fill != null)
            {
                var fillValue = ReadFillValue(fill);
                style.Pattern = fillValue.Pattern;
                style.ForegroundColor = fillValue.ForegroundColor;
                style.BackgroundColor = fillValue.BackgroundColor;
            }
            var border = dxf.Element(MainNs + "border");
            if (border != null)
            {
                style.Borders = ReadBordersValue(border);
            }
            var numFmt = dxf.Element(MainNs + "numFmt");
            if (numFmt != null)
            {
                style.NumberFormat = new NumberFormatValue
                {
                    Number = ParseIntAttribute(numFmt.Attribute("numFmtId")) ?? 0,
                    Custom = (string)numFmt.Attribute("formatCode"),
                };
            }
            var alignment = dxf.Element(MainNs + "alignment");
            if (alignment != null)
            {
                style.Alignment = new AlignmentValue
                {
                    Horizontal = ParseHorizontalAlignment((string)alignment.Attribute("horizontal")),
                    Vertical = ParseVerticalAlignment((string)alignment.Attribute("vertical")),
                    WrapText = ParseBoolAttribute(alignment.Attribute("wrapText")),
                    IndentLevel = StyleValueSanitizer.NormalizeIndentLevel(ParseIntAttribute(alignment.Attribute("indent"))),
                    TextRotation = StyleValueSanitizer.NormalizeTextRotation(ParseIntAttribute(alignment.Attribute("textRotation"))),
                    ShrinkToFit = ParseBoolAttribute(alignment.Attribute("shrinkToFit")),
                    ReadingOrder = StyleValueSanitizer.NormalizeReadingOrder(ParseIntAttribute(alignment.Attribute("readingOrder"))),
                    RelativeIndent = ParseIntAttribute(alignment.Attribute("relativeIndent")) ?? 0,
                };
            }
            var protection = dxf.Element(MainNs + "protection");
            if (protection != null)
            {
                style.Protection = new ProtectionValue
                {
                    IsLocked = ParseOptionalBoolAttribute(protection.Attribute("locked")) ?? true,
                    IsHidden = ParseBoolAttribute(protection.Attribute("hidden")),
                };
            }
            return style;
        }
        private static FontValue ReadFontValue(XElement font)
        {
            return new FontValue
            {
                Name = (string)font.Element(MainNs + "name")?.Attribute("val") ?? "Calibri",
                Size = ParseDoubleAttribute(font.Element(MainNs + "sz")?.Attribute("val")) ?? 11d,
                Bold = font.Element(MainNs + "b") != null,
                Italic = font.Element(MainNs + "i") != null,
                Underline = font.Element(MainNs + "u") != null,
                StrikeThrough = font.Element(MainNs + "strike") != null,
                Color = ReadColorValue(font.Element(MainNs + "color")),
            };
        }
        private static FillValue ReadFillValue(XElement fill)
        {
            var patternFill = fill.Element(MainNs + "patternFill");
            if (patternFill == null)
            {
                return new FillValue();
            }
            var patternType = ((string)patternFill.Attribute("patternType") ?? "none").ToLowerInvariant();
            FillPatternKind pattern;
            switch (patternType)
            {
                case "solid":
                    pattern = FillPatternKind.Solid;
                    break;
                case "mediumgray":
                    pattern = FillPatternKind.MediumGray;
                    break;
                case "darkgray":
                    pattern = FillPatternKind.DarkGray;
                    break;
                case "gray125":
                    pattern = FillPatternKind.Gray125;
                    break;
                case "gray0625":
                    pattern = FillPatternKind.Gray0625;
                    break;
                case "darkhorizontal":
                    pattern = FillPatternKind.DarkHorizontal;
                    break;
                case "darkvertical":
                    pattern = FillPatternKind.DarkVertical;
                    break;
                case "darkdown":
                    pattern = FillPatternKind.DarkDown;
                    break;
                case "darkup":
                    pattern = FillPatternKind.DarkUp;
                    break;
                case "darkgrid":
                    pattern = FillPatternKind.DarkGrid;
                    break;
                case "darktrellis":
                    pattern = FillPatternKind.DarkTrellis;
                    break;
                case "lighthorizontal":
                    pattern = FillPatternKind.LightHorizontal;
                    break;
                case "lightvertical":
                    pattern = FillPatternKind.LightVertical;
                    break;
                case "lightdown":
                    pattern = FillPatternKind.LightDown;
                    break;
                case "lightup":
                    pattern = FillPatternKind.LightUp;
                    break;
                case "lightgrid":
                    pattern = FillPatternKind.LightGrid;
                    break;
                case "lighttrellis":
                    pattern = FillPatternKind.LightTrellis;
                    break;
                default:
                    pattern = FillPatternKind.None;
                    break;
            }
            return new FillValue
            {
                Pattern = pattern,
                ForegroundColor = ReadColorValue(patternFill.Element(MainNs + "fgColor")),
                BackgroundColor = ReadColorValue(patternFill.Element(MainNs + "bgColor")),
            };
        }
        private static BordersValue ReadBordersValue(XElement border)
        {
            return new BordersValue
            {
                Left = ReadBorderSideValue(border.Element(MainNs + "left")),
                Right = ReadBorderSideValue(border.Element(MainNs + "right")),
                Top = ReadBorderSideValue(border.Element(MainNs + "top")),
                Bottom = ReadBorderSideValue(border.Element(MainNs + "bottom")),
                Diagonal = ReadBorderSideValue(border.Element(MainNs + "diagonal")),
                DiagonalUp = ParseBoolAttribute(border.Attribute("diagonalUp")),
                DiagonalDown = ParseBoolAttribute(border.Attribute("diagonalDown")),
            };
        }
        private static XElement BuildFontElement(FontValue font)
        {
            var element = new XElement(MainNs + "font");
            if (font.Bold)
            {
                element.Add(new XElement(MainNs + "b"));
            }
            if (font.Italic)
            {
                element.Add(new XElement(MainNs + "i"));
            }
            if (font.StrikeThrough)
            {
                element.Add(new XElement(MainNs + "strike"));
            }
            if (font.Underline)
            {
                element.Add(new XElement(MainNs + "u"));
            }
            element.Add(new XElement(MainNs + "sz", new XAttribute("val", font.Size.ToString("0.####", CultureInfo.InvariantCulture))));
            var colorElement = BuildColorElement("color", font.Color);
            if (colorElement != null)
            {
                element.Add(colorElement);
            }
            element.Add(new XElement(MainNs + "name", new XAttribute("val", font.Name)));
            return element;
        }
        private static XElement BuildFillElement(FillValue fill)
        {
            var patternFill = new XElement(MainNs + "patternFill");
            switch (fill.Pattern)
            {
                case FillPatternKind.Solid:
                    patternFill.SetAttributeValue("patternType", "solid");
                    break;
                case FillPatternKind.MediumGray:
                    patternFill.SetAttributeValue("patternType", "mediumGray");
                    break;
                case FillPatternKind.DarkGray:
                    patternFill.SetAttributeValue("patternType", "darkGray");
                    break;
                case FillPatternKind.Gray125:
                    patternFill.SetAttributeValue("patternType", "gray125");
                    break;
                case FillPatternKind.Gray0625:
                    patternFill.SetAttributeValue("patternType", "gray0625");
                    break;
                case FillPatternKind.DarkHorizontal:
                    patternFill.SetAttributeValue("patternType", "darkHorizontal");
                    break;
                case FillPatternKind.DarkVertical:
                    patternFill.SetAttributeValue("patternType", "darkVertical");
                    break;
                case FillPatternKind.DarkDown:
                    patternFill.SetAttributeValue("patternType", "darkDown");
                    break;
                case FillPatternKind.DarkUp:
                    patternFill.SetAttributeValue("patternType", "darkUp");
                    break;
                case FillPatternKind.DarkGrid:
                    patternFill.SetAttributeValue("patternType", "darkGrid");
                    break;
                case FillPatternKind.DarkTrellis:
                    patternFill.SetAttributeValue("patternType", "darkTrellis");
                    break;
                case FillPatternKind.LightHorizontal:
                    patternFill.SetAttributeValue("patternType", "lightHorizontal");
                    break;
                case FillPatternKind.LightVertical:
                    patternFill.SetAttributeValue("patternType", "lightVertical");
                    break;
                case FillPatternKind.LightDown:
                    patternFill.SetAttributeValue("patternType", "lightDown");
                    break;
                case FillPatternKind.LightUp:
                    patternFill.SetAttributeValue("patternType", "lightUp");
                    break;
                case FillPatternKind.LightGrid:
                    patternFill.SetAttributeValue("patternType", "lightGrid");
                    break;
                case FillPatternKind.LightTrellis:
                    patternFill.SetAttributeValue("patternType", "lightTrellis");
                    break;
                default:
                    patternFill.SetAttributeValue("patternType", "none");
                    break;
            }
            var foregroundColor = BuildColorElement("fgColor", fill.ForegroundColor);
            if (foregroundColor != null)
            {
                patternFill.Add(foregroundColor);
            }
            var backgroundColor = BuildColorElement("bgColor", fill.BackgroundColor);
            if (backgroundColor != null)
            {
                patternFill.Add(backgroundColor);
            }
            return new XElement(MainNs + "fill", patternFill);
        }
        private static XElement BuildBorderElement(BordersValue borders)
        {
            var element = new XElement(MainNs + "border",
                BuildBorderSideElement("left", borders.Left),
                BuildBorderSideElement("right", borders.Right),
                BuildBorderSideElement("top", borders.Top),
                BuildBorderSideElement("bottom", borders.Bottom),
                BuildBorderSideElement("diagonal", borders.Diagonal));
            if (borders.DiagonalUp)
            {
                element.SetAttributeValue("diagonalUp", 1);
            }
            if (borders.DiagonalDown)
            {
                element.SetAttributeValue("diagonalDown", 1);
            }
            return element;
        }
        private static XElement BuildCellFormatElement(CellFormatValue cellFormat)
        {
            return BuildCellFormatElement(cellFormat, true);
        }
        private static XElement BuildDifferentialFormatElement(StyleValue style)
        {
            var element = new XElement(MainNs + "dxf");
            if (!FontEquals(style.Font, StyleValue.Default.Font))
            {
                element.Add(BuildFontElement(style.Font));
            }
            if (style.Pattern != FillPatternKind.None || !IsEmptyColor(style.ForegroundColor) || !IsEmptyColor(style.BackgroundColor))
            {
                element.Add(BuildFillElement(new FillValue
                {
                    Pattern = style.Pattern,
                    ForegroundColor = style.ForegroundColor,
                    BackgroundColor = style.BackgroundColor,
                }));
            }
            if (!BordersEqual(style.Borders, StyleValue.Default.Borders))
            {
                element.Add(BuildBorderElement(style.Borders));
            }
            if (style.NumberFormat.Number != 0 || !string.IsNullOrEmpty(style.NumberFormat.Custom))
            {
                var numFmtElement = new XElement(MainNs + "numFmt",
                    new XAttribute("numFmtId", style.NumberFormat.Number >= 0 ? style.NumberFormat.Number : 0));
                if (!string.IsNullOrEmpty(style.NumberFormat.Custom))
                {
                    numFmtElement.SetAttributeValue("formatCode", style.NumberFormat.Custom);
                }
                element.Add(numFmtElement);
            }
            var alignmentElement = BuildAlignmentElement(style.Alignment);
            if (alignmentElement != null)
            {
                element.Add(alignmentElement);
            }
            var protectionElement = BuildProtectionElement(style.Protection);
            if (protectionElement != null)
            {
                element.Add(protectionElement);
            }
            return element;
        }
        private static XElement BuildCellFormatElement(CellFormatValue cellFormat, bool includeXfId)
        {
            var element = new XElement(MainNs + "xf",
                new XAttribute("numFmtId", cellFormat.NumFmtId),
                new XAttribute("fontId", cellFormat.FontId),
                new XAttribute("fillId", cellFormat.FillId),
                new XAttribute("borderId", cellFormat.BorderId));
            if (includeXfId)
            {
                element.SetAttributeValue("xfId", 0);
            }
            if (cellFormat.NumFmtId != 0)
            {
                element.SetAttributeValue("applyNumberFormat", 1);
            }
            if (cellFormat.FontId != 0)
            {
                element.SetAttributeValue("applyFont", 1);
            }
            if (cellFormat.FillId != 0)
            {
                element.SetAttributeValue("applyFill", 1);
            }
            if (cellFormat.BorderId != 0)
            {
                element.SetAttributeValue("applyBorder", 1);
            }
            var alignmentElement = BuildAlignmentElement(cellFormat.Alignment);
            if (alignmentElement != null)
            {
                element.SetAttributeValue("applyAlignment", 1);
                element.Add(alignmentElement);
            }
            var protectionElement = BuildProtectionElement(cellFormat.Protection);
            if (protectionElement != null)
            {
                element.SetAttributeValue("applyProtection", 1);
                element.Add(protectionElement);
            }
            return element;
        }
        private static BorderSideValue ReadBorderSideValue(XElement side)
        {
            if (side == null)
            {
                return new BorderSideValue();
            }
            return new BorderSideValue
            {
                Style = ParseBorderStyle((string)side.Attribute("style")),
                Color = ReadColorValue(side.Element(MainNs + "color")),
            };
        }
        private static ColorValue ReadColorValue(XElement colorElement)
        {
            if (colorElement == null)
            {
                return default(ColorValue);
            }
            var rgb = (string)colorElement.Attribute("rgb");
            if (string.IsNullOrWhiteSpace(rgb))
            {
                return default(ColorValue);
            }
            var rgbValue = rgb.Trim().ToUpperInvariant();
            if (rgbValue.Length == 6)
            {
                rgbValue = "FF" + rgbValue;
            }
            if (rgbValue.Length != 8)
            {
                return default(ColorValue);
            }
            byte a, r, g, b;
            if (!byte.TryParse(rgbValue.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out a))
            {
                return default(ColorValue);
            }
            if (!byte.TryParse(rgbValue.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r))
            {
                return default(ColorValue);
            }
            if (!byte.TryParse(rgbValue.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g))
            {
                return default(ColorValue);
            }
            if (!byte.TryParse(rgbValue.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b))
            {
                return default(ColorValue);
            }
            return new ColorValue(a, r, g, b);
        }
        private static XElement BuildColorElement(string elementName, ColorValue color)
        {
            if (IsEmptyColor(color))
            {
                return null;
            }
            return new XElement(MainNs + elementName, new XAttribute("rgb", ToArgbHex(color)));
        }
        private static XElement BuildBorderSideElement(string sideName, BorderSideValue side)
        {
            var element = new XElement(MainNs + sideName);
            var styleName = GetBorderStyleName(side.Style);
            if (!string.IsNullOrEmpty(styleName))
            {
                element.SetAttributeValue("style", styleName);
            }
            var colorElement = BuildColorElement("color", side.Color);
            if (colorElement != null)
            {
                element.Add(colorElement);
            }
            return element;
        }
    
        private static XElement BuildAlignmentElement(AlignmentValue alignment)
        {
            var element = new XElement(MainNs + "alignment");
            var hasValue = false;
            var horizontal = GetHorizontalAlignmentName(alignment.Horizontal);
            if (!string.IsNullOrEmpty(horizontal))
            {
                element.SetAttributeValue("horizontal", horizontal);
                hasValue = true;
            }
            var vertical = GetVerticalAlignmentName(alignment.Vertical);
            if (!string.IsNullOrEmpty(vertical))
            {
                element.SetAttributeValue("vertical", vertical);
                hasValue = true;
            }
            if (alignment.WrapText)
            {
                element.SetAttributeValue("wrapText", 1);
                hasValue = true;
            }
            if (alignment.IndentLevel > 0)
            {
                element.SetAttributeValue("indent", alignment.IndentLevel);
                hasValue = true;
            }
            if (alignment.TextRotation != 0)
            {
                element.SetAttributeValue("textRotation", alignment.TextRotation);
                hasValue = true;
            }
            if (alignment.ShrinkToFit)
            {
                element.SetAttributeValue("shrinkToFit", 1);
                hasValue = true;
            }
            if (alignment.ReadingOrder != 0)
            {
                element.SetAttributeValue("readingOrder", alignment.ReadingOrder);
                hasValue = true;
            }
            if (alignment.RelativeIndent != 0)
            {
                element.SetAttributeValue("relativeIndent", alignment.RelativeIndent);
                hasValue = true;
            }
            return hasValue ? element : null;
        }
    
        private static XElement BuildProtectionElement(ProtectionValue protection)
        {
            var element = new XElement(MainNs + "protection");
            var hasValue = false;
            if (!protection.IsLocked)
            {
                element.SetAttributeValue("locked", 0);
                hasValue = true;
            }
            if (protection.IsHidden)
            {
                element.SetAttributeValue("hidden", 1);
                hasValue = true;
            }
            return hasValue ? element : null;
        }
    
    
    
    
    
    }
}
