namespace Aspose.Cells_FOSS;

public static class NumberFormat
{
    private static readonly Dictionary<int, string> BuiltInFormats = CreateBuiltInFormats();

    public static string GetBuiltInFormat(int formatId)
    {
        if (BuiltInFormats.TryGetValue(formatId, out var formatCode))
        {
            return formatCode;
        }

        return "General";
    }

    public static bool IsBuiltInFormat(string formatCode)
    {
        return GetBuiltInFormatId(formatCode).HasValue;
    }

    public static int? GetBuiltInFormatId(string? formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode))
        {
            return 0;
        }

        var trimmed = formatCode is null ? string.Empty : formatCode.Trim();
        foreach (var pair in BuiltInFormats)
        {
            if (string.Equals(pair.Value, trimmed, StringComparison.Ordinal))
            {
                return pair.Key;
            }
        }

        return null;
    }

    internal static string ResolveFormatCode(int number, string? custom)
    {
        if (!string.IsNullOrEmpty(custom))
        {
            return custom!;
        }

        return GetBuiltInFormat(number);
    }

    private static Dictionary<int, string> CreateBuiltInFormats()
    {
        return new Dictionary<int, string>
        {
            [0] = "General",
            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [4] = "#,##0.00",
            [5] = "$#,##0_);($#,##0)",
            [6] = "$#,##0_);[Red]($#,##0)",
            [7] = "$#,##0.00_);($#,##0.00)",
            [8] = "$#,##0.00_);[Red]($#,##0.00)",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00",
            [12] = "# ?/?",
            [13] = "# ??/??",
            [14] = "mm-dd-yy",
            [15] = "d-mmm-yy",
            [16] = "d-mmm",
            [17] = "mmm-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "m/d/yy h:mm",
            [37] = "#,##0_);(#,##0)",
            [38] = "#,##0_);[Red](#,##0)",
            [39] = "#,##0.00_);(#,##0.00)",
            [40] = "#,##0.00_);[Red](#,##0.00)",
            [41] = "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
            [42] = "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
            [43] = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
            [44] = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mm:ss.0",
            [48] = "##0.0E+0",
            [49] = "@",
            [50] = "General",
            [51] = "0_);(0)",
            [52] = "0_);[Red](0)",
            [53] = "0_);(0)",
            [54] = "0_);[Red](0)",
            [55] = "0_);(0)",
            [56] = "0_);[Red](0)",
            [57] = "0_);(0)",
            [58] = "0_);[Red](0)",
            [59] = "0_);(0)",
            [60] = "0_);[Red](0)",
            [61] = "0_);(0)",
            [62] = "0_);[Red](0)",
            [63] = "0_);(0)",
            [64] = "0_);[Red](0)",
            [65] = "0_);(0)",
            [66] = "0_);[Red](0)",
            [67] = "0_);(0)",
            [68] = "0_);[Red](0)",
            [69] = "0_);(0)",
            [70] = "0_);[Red](0)",
            [71] = "0_);(0)",
            [72] = "0_);[Red](0)",
            [73] = "0_);(0)",
            [74] = "0_);[Red](0)",
            [75] = "0_);(0)",
            [76] = "0_);[Red](0)",
            [77] = "0_);(0)",
            [78] = "0_);[Red](0)",
            [79] = "0_);(0)",
            [80] = "0_);[Red](0)",
            [81] = "0_);(0)",
            [82] = "0_);[Red](0)",
        };
    }
}

