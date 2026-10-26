using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Legacy Office chart styles 2/3 supply line defaults when no modern style part exists.</summary>
    internal static class ChartLegacyGridStyle
    {
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public static void Apply(XElement root, XElement plot, ChartModel model, ParsedChart chart)
        {
            if (root.Element(Mc + "AlternateContent") != null) return;
            if (model != null)
                foreach (var companion in model.CompanionFiles)
                    if (companion.RelationshipType == "http://schemas.microsoft.com/office/2011/relationships/chartStyle") return;
            var style = (string)root.Element(C + "style")?.Attribute("val");
            if (style != null && style != "2" && style != "3") return;
            var axis = plot.Element(C + "valAx");
            var grid = axis?.Element(C + "majorGridlines");
            var gridLine = grid?.Element(C + "spPr")?.Element(A + "ln");
            if (grid != null)
            {
                if (gridLine?.Attribute("w") == null) chart.ValueGridlineWidthPt = 0.75d;
                if (gridLine == null) chart.ValueGridlineColor = new SKColor(0x86, 0x86, 0x86);
            }
            if (axis == null) return;
            var deleted = (string)axis.Element(C + "delete")?.Attribute("val");
            if (deleted == "1" || deleted == "true")
            {
                chart.HasValueAxisLine = false;
                return;
            }
            // Missing line properties inherit the legacy style; explicit noFill does not.
            if (axis.Element(C + "spPr")?.Element(A + "ln") == null)
            {
                chart.HasValueAxisLine = true;
                chart.ValueAxisLineColor = new SKColor(0x86, 0x86, 0x86);
                chart.ValueAxisLineWidthPt = 0.75d;
            }
            if (chart.ValueAxisMajorTickMark == null) chart.ValueAxisMajorTickMark = "out";
        }
    }
}
