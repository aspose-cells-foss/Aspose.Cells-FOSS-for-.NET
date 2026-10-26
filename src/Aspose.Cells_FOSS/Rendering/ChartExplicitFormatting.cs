using System;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Inherited chart text and explicit point formatting, independent of chart data.</summary>
    internal static class ChartExplicitFormatting
    {
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static void ApplyTextDefaults(XElement root, RenderColor colors, ParsedChart chart)
        {
            var properties = root.Element(C + "txPr")?.Element(A + "p")?.Element(A + "pPr")?.Element(A + "defRPr");
            var size = ReadSize(properties);
            if (!size.HasValue) return;
            chart.TextSizePt = size;
            chart.TitleFont = new FontValue { Name = colors.MinorFontName ?? "Calibri", Size = size.Value * 1.2d, Bold = true };
            var color = colors.ResolveSchemeName("tx1", SKColors.Black);
            SKColor explicitColor;
            if (ChartXmlParser.TryDrawingColor(ChartXmlParser.FillChild(properties.Element(A + "solidFill")), colors, out explicitColor))
            {
                color = explicitColor;
            }
            chart.TitleColor = color;
            chart.AxisLabelColor = color;
            chart.LegendTextColor = color;
        }

        public static void ApplySeries(XElement element, XElement type, RenderColor colors, ChartSeries series, int fallbackIndex)
        {
            int index;
            if (!int.TryParse((string)element.Element(C + "idx")?.Attribute("val"), out index) || index < 0) index = fallbackIndex;
            var shape = element.Element(C + "spPr");
            if (shape?.Element(A + "solidFill") == null && shape?.Element(A + "ln")?.Element(A + "solidFill") == null)
            {
                series.Color = AutomaticColor(colors, index);
            }
            if ((string)type.Element(C + "varyColors")?.Attribute("val") == "1")
            {
                for (var i = 0; i < series.Values.Count; i++) series.PointColors[i] = AutomaticColor(colors, i);
            }
            foreach (var point in element.Elements(C + "dPt"))
            {
                int pointIndex;
                SKColor color;
                if (int.TryParse((string)point.Element(C + "idx")?.Attribute("val"), out pointIndex) && pointIndex >= 0
                    && ChartXmlParser.TryDrawingColor(ChartXmlParser.FillChild(point.Element(C + "spPr")?.Element(A + "solidFill")), colors, out color))
                {
                    series.PointColors[pointIndex] = color;
                }
            }
            var labels = element.Element(C + "dLbls") ?? type.Element(C + "dLbls");
            series.DataLabelSizePt = ReadSize(labels?.Element(C + "txPr")?.Element(A + "p")?.Element(A + "pPr")?.Element(A + "defRPr"));
        }

        public static SKColor PointColor(ChartSeries series, int index)
        {
            SKColor color;
            return series.PointColors.TryGetValue(index, out color) ? color : series.Color;
        }

        private static SKColor AutomaticColor(RenderColor colors, int index)
        {
            var color = colors.ResolveSchemeName("accent" + (index % 6 + 1), SKColors.Gray);
            if (index >= 6) color = RenderColor.ApplyLuma(color, 0.6d, 0.4d);
            return color;
        }

        private static double? ReadSize(XElement properties)
        {
            double size;
            if (double.TryParse((string)properties?.Attribute("sz"), NumberStyles.Float, CultureInfo.InvariantCulture, out size) && size > 0d)
                return size / 100d;
            return null;
        }
    }
}
