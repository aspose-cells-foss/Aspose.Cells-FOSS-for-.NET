using System;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Explicit category-label lines share the same font metrics in layout and drawing.</summary>
    internal static class ChartCategoryLabels
    {
        public static string[] Lines(string label)
        {
            return (label ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        }

        public static float LineAdvance(CellFontContext font)
        {
            return Math.Max(font.LineHeightPt, font.TextHeightPt);
        }

        public static float ExtraHeight(ParsedChart chart, CellFontContext font)
        {
            var lines = 1;
            foreach (var label in chart.Categories)
                lines = Math.Max(lines, Lines(label).Length);
            return (lines - 1) * LineAdvance(font);
        }
    }
}
