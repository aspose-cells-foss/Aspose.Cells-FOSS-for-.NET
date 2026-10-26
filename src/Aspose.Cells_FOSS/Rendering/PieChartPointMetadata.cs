using System;
using System.Globalization;
using System.Xml.Linq;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class PieChartPointMetadata
    {
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static void Apply(XElement series, RenderColor colors, ParsedChart chart)
        {
            if (chart == null || chart.Kind != ChartKind.Pie || chart.Series.Count == 0)
            {
                return;
            }

            var pointCount = chart.Series[0].Values.Count;
            if (chart.Categories.Count == 0)
            {
                for (var i = 0; i < pointCount; i++)
                {
                    chart.Categories.Add((i + 1).ToString(CultureInfo.InvariantCulture));
                }
            }

            for (var i = 0; i < pointCount; i++)
            {
                chart.PiePointLineColors.Add(SKColors.Transparent);
                chart.PiePointLineWidthsPt.Add(0d);
            }

            if (series == null)
            {
                return;
            }

            foreach (var point in series.Elements(C + "dPt"))
            {
                var indexElement = point.Element(C + "idx");
                int index;
                if (indexElement == null
                    || !int.TryParse((string)indexElement.Attribute("val"), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out index)
                    || index < 0 || index >= pointCount)
                {
                    continue;
                }

                var shape = point.Element(C + "spPr");
                var line = shape != null ? shape.Element(A + "ln") : null;
                bool visible;
                SKColor color;
                double width;
                ChartXmlParser.ParseLineStyle(line, colors, out visible, out color, out width);
                if (visible)
                {
                    chart.PiePointLineColors[index] = color;
                    chart.PiePointLineWidthsPt[index] = width;
                }
            }
        }
    }
}
