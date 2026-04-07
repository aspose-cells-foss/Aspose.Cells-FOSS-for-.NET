using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

internal static class XlsxWorkbookStylesValueHelpers
{
    internal static BorderStyle ParseBorderStyle(string? value)
    {
        switch (value?.ToLowerInvariant())
        {
            case "thin":
                return BorderStyle.Thin;
            case "medium":
                return BorderStyle.Medium;
            case "thick":
                return BorderStyle.Thick;
            case "dotted":
                return BorderStyle.Dotted;
            case "dashed":
                return BorderStyle.Dashed;
            case "double":
                return BorderStyle.Double;
            case "hair":
                return BorderStyle.Hair;
            case "mediumdashed":
                return BorderStyle.MediumDashed;
            case "dashdot":
                return BorderStyle.DashDot;
            case "mediumdashdot":
                return BorderStyle.MediumDashDot;
            case "dashdotdot":
                return BorderStyle.DashDotDot;
            case "mediumdashdotdot":
                return BorderStyle.MediumDashDotDot;
            case "slantdashdot":
                return BorderStyle.SlantedDashDot;
            default:
                return BorderStyle.None;
        }
    }

    internal static string GetBorderStyleName(BorderStyle value)
    {
        switch (value)
        {
            case BorderStyle.Thin:
                return "thin";
            case BorderStyle.Medium:
                return "medium";
            case BorderStyle.Thick:
                return "thick";
            case BorderStyle.Dotted:
                return "dotted";
            case BorderStyle.Dashed:
                return "dashed";
            case BorderStyle.Double:
                return "double";
            case BorderStyle.Hair:
                return "hair";
            case BorderStyle.MediumDashed:
                return "mediumDashed";
            case BorderStyle.DashDot:
                return "dashDot";
            case BorderStyle.MediumDashDot:
                return "mediumDashDot";
            case BorderStyle.DashDotDot:
                return "dashDotDot";
            case BorderStyle.MediumDashDotDot:
                return "mediumDashDotDot";
            case BorderStyle.SlantedDashDot:
                return "slantDashDot";
            default:
                return string.Empty;
        }
    }

    internal static string GetHorizontalAlignmentName(HorizontalAlignment value)
    {
        switch (value)
        {
            case HorizontalAlignment.Left:
                return "left";
            case HorizontalAlignment.Center:
                return "center";
            case HorizontalAlignment.Right:
                return "right";
            case HorizontalAlignment.Fill:
                return "fill";
            case HorizontalAlignment.Justify:
                return "justify";
            case HorizontalAlignment.CenterContinuous:
                return "centerContinuous";
            case HorizontalAlignment.Distributed:
                return "distributed";
            default:
                return string.Empty;
        }
    }

    internal static string GetVerticalAlignmentName(VerticalAlignment value)
    {
        switch (value)
        {
            case VerticalAlignment.Center:
                return "center";
            case VerticalAlignment.Top:
                return "top";
            case VerticalAlignment.Justify:
                return "justify";
            case VerticalAlignment.Distributed:
                return "distributed";
            default:
                return string.Empty;
        }
    }

    internal static string ToArgbHex(ColorValue color)
    {
        return string.Concat(
            color.A.ToString("X2", CultureInfo.InvariantCulture),
            color.R.ToString("X2", CultureInfo.InvariantCulture),
            color.G.ToString("X2", CultureInfo.InvariantCulture),
            color.B.ToString("X2", CultureInfo.InvariantCulture));
    }

    internal static bool IsEmptyColor(ColorValue color)
    {
        return color.A == 0 && color.R == 0 && color.G == 0 && color.B == 0;
    }

    internal static bool FontEquals(FontValue left, FontValue right)
    {
        return string.Equals(left.Name, right.Name, StringComparison.Ordinal)
            && left.Size.Equals(right.Size)
            && left.Bold == right.Bold
            && left.Italic == right.Italic
            && left.Underline == right.Underline
            && left.StrikeThrough == right.StrikeThrough
            && left.Color.Equals(right.Color);
    }

    internal static bool BordersEqual(BordersValue left, BordersValue right)
    {
        return BorderSideEquals(left.Left, right.Left)
            && BorderSideEquals(left.Right, right.Right)
            && BorderSideEquals(left.Top, right.Top)
            && BorderSideEquals(left.Bottom, right.Bottom)
            && BorderSideEquals(left.Diagonal, right.Diagonal)
            && left.DiagonalUp == right.DiagonalUp
            && left.DiagonalDown == right.DiagonalDown;
    }

    private static bool BorderSideEquals(BorderSideValue left, BorderSideValue right)
    {
        return left.Style == right.Style && left.Color.Equals(right.Color);
    }
}
