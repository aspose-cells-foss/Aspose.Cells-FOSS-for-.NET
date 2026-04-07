using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.CompareOpenXml;

internal static class ComparisonValueHelpers
{
    internal const string DefaultColorHex = "00000000";
    internal static readonly XNamespace SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal static readonly HashSet<int> BuiltInDateFormats = new HashSet<int> { 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47 };

    internal static bool TryNormalizeCellReference(string? cellReference, out string normalized)
    {
        normalized = string.Empty;
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return false;
        }

        try
        {
            normalized = CellAddress.Parse(cellReference).ToString();
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    internal static string BuildCellKey(string sheetName, string cellReference)
    {
        return sheetName + "!" + cellReference.ToUpperInvariant();
    }

    internal static string NormalizeValue(object? value)
    {
        if (value is null)
        {
            return "<null>";
        }

        if (value is string text)
        {
            return text;
        }

        if (value is bool booleanValue)
        {
            return booleanValue ? "TRUE" : "FALSE";
        }

        if (value is DateTime dateTime)
        {
            return dateTime.ToString("O", CultureInfo.InvariantCulture);
        }

        if (value is IFormattable formattable)
        {
            return formattable.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        return value.ToString() ?? string.Empty;
    }

    internal static string NormalizeFormulaText(string? formulaBody)
    {
        if (string.IsNullOrWhiteSpace(formulaBody))
        {
            return string.Empty;
        }

        var normalized = formulaBody.Trim();
        return normalized.StartsWith("=", StringComparison.Ordinal) ? normalized : "=" + normalized;
    }

    internal static string NormalizeDouble(double value)
    {
        return value.ToString("0.####", CultureInfo.InvariantCulture);
    }

    internal static string NormalizeColor(Color color)
    {
        return color.Equals(Color.Empty)
            ? DefaultColorHex
            : string.Concat(
                color.A.ToString("X2", CultureInfo.InvariantCulture),
                color.R.ToString("X2", CultureInfo.InvariantCulture),
                color.G.ToString("X2", CultureInfo.InvariantCulture),
                color.B.ToString("X2", CultureInfo.InvariantCulture));
    }

    internal static string NormalizeOpenXmlColor(XElement? colorElement)
    {
        if (colorElement is null)
        {
            return DefaultColorHex;
        }

        var rgb = (string?)colorElement.Attribute("rgb");
        if (string.IsNullOrWhiteSpace(rgb))
        {
            return DefaultColorHex;
        }

        rgb = rgb.Trim().ToUpperInvariant();
        if (rgb.Length == 6)
        {
            rgb = "FF" + rgb;
        }

        return rgb.Length == 8 ? rgb : DefaultColorHex;
    }

    internal static string ParseHorizontalAlignment(string? value)
    {
        switch (value?.ToLowerInvariant())
        {
            case "left":
                return "Left";
            case "center":
                return "Center";
            case "right":
                return "Right";
            case "fill":
                return "Fill";
            case "justify":
                return "Justify";
            case "centercontinuous":
                return "CenterContinuous";
            case "distributed":
                return "Distributed";
            default:
                return "General";
        }
    }

    internal static string ParseVerticalAlignment(string? value)
    {
        switch (value?.ToLowerInvariant())
        {
            case "center":
                return "Center";
            case "top":
                return "Top";
            case "justify":
                return "Justify";
            case "distributed":
                return "Distributed";
            default:
                return "Bottom";
        }
    }

    internal static string ParseBorderStyle(string? value)
    {
        switch (value?.ToLowerInvariant())
        {
            case "thin":
                return "Thin";
            case "medium":
                return "Medium";
            case "thick":
                return "Thick";
            case "dotted":
                return "Dotted";
            case "dashed":
                return "Dashed";
            case "double":
                return "Double";
            case "hair":
                return "Hair";
            case "mediumdashed":
                return "MediumDashed";
            case "dashdot":
                return "DashDot";
            case "mediumdashdot":
                return "MediumDashDot";
            case "dashdotdot":
                return "DashDotDot";
            case "mediumdashdotdot":
                return "MediumDashDotDot";
            case "slantdashdot":
                return "SlantedDashDot";
            default:
                return "None";
        }
    }

    internal static string ParseFillPattern(string? value)
    {
        switch (value?.ToLowerInvariant())
        {
            case "solid":
                return "Solid";
            case "mediumgray":
                return "MediumGray";
            case "darkgray":
                return "DarkGray";
            case "gray125":
                return "Gray125";
            case "gray0625":
                return "Gray0625";
            case "darkhorizontal":
                return "DarkHorizontal";
            case "darkvertical":
                return "DarkVertical";
            case "darkdown":
                return "DarkDown";
            case "darkup":
                return "DarkUp";
            case "darkgrid":
                return "DarkGrid";
            case "darktrellis":
                return "DarkTrellis";
            case "lighthorizontal":
                return "LightHorizontal";
            case "lightvertical":
                return "LightVertical";
            case "lightdown":
                return "LightDown";
            case "lightup":
                return "LightUp";
            case "lightgrid":
                return "LightGrid";
            case "lighttrellis":
                return "LightTrellis";
            default:
                return "None";
        }
    }

    internal static bool ParseBoolAttribute(XAttribute? attribute)
    {
        if (attribute is null)
        {
            return false;
        }

        var value = attribute.Value;
        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    internal static bool? ParseOptionalBoolAttribute(XAttribute? attribute)
    {
        return attribute is null ? null : ParseBoolAttribute(attribute);
    }

    internal static int? ParseIntAttribute(XAttribute? attribute)
    {
        if (attribute is null)
        {
            return null;
        }

        return int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : null;
    }

    internal static double ParseDoubleValue(string? value, double defaultValue)
    {
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : defaultValue;
    }

    internal static bool TryParseNumber(string rawValue, out object? value, out string kind)
    {
        if (!rawValue.Contains('e') && !rawValue.Contains('E') && decimal.TryParse(rawValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
        {
            kind = "Number";
            if (!rawValue.Contains('.') && !rawValue.Contains(',') && decimal.Truncate(decimalValue) == decimalValue && decimalValue >= int.MinValue && decimalValue <= int.MaxValue)
            {
                value = (int)decimalValue;
                return true;
            }

            value = decimalValue;
            return true;
        }

        if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var doubleValue))
        {
            kind = "Number";
            if (Math.Abs(doubleValue % 1d) < double.Epsilon && doubleValue >= int.MinValue && doubleValue <= int.MaxValue)
            {
                value = (int)doubleValue;
                return true;
            }

            value = doubleValue;
            return true;
        }

        kind = "Blank";
        value = null;
        return false;
    }

    internal static bool LooksLikeDateFormat(string formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode))
        {
            return false;
        }

        var builder = new System.Text.StringBuilder(formatCode.Length);
        var inQuote = false;
        var inBracket = false;

        for (var index = 0; index < formatCode.Length; index++)
        {
            var character = formatCode[index];
            if (character == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (inQuote)
            {
                continue;
            }

            if (character == '[')
            {
                inBracket = true;
                continue;
            }

            if (character == ']' && inBracket)
            {
                inBracket = false;
                continue;
            }

            if (inBracket)
            {
                builder.Append(char.ToLowerInvariant(character));
                continue;
            }

            if (character == '\\' || character == '_')
            {
                index++;
                continue;
            }

            if (character == '*')
            {
                continue;
            }

            builder.Append(char.ToLowerInvariant(character));
        }

        var normalized = builder.ToString();
        return normalized.IndexOf("yy", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("dd", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("mm", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("m/", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("/m", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("h", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("ss", StringComparison.Ordinal) >= 0;
    }

    internal static T GetOrDefault<T>(IReadOnlyList<T> values, int index, T fallback)
    {
        return index >= 0 && index < values.Count ? values[index] : fallback;
    }
}
