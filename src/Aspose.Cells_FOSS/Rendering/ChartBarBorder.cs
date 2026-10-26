using System.Xml.Linq;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Explicit outlines for 2-D columns and bars, separate from line-series defaults.</summary>
    internal static class ChartBarBorder
    {
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static void Apply(XElement element, ChartSeries series)
        {
            var line = element.Element(C + "spPr")?.Element(A + "ln");
            series.HasBarBorder = line != null && series.HasVisibleLine;
            if (line != null && line.Attribute("w") == null) series.LineWidthPt = 0.75d;
        }

        public static void Draw(SKCanvas canvas, SKRect bounds, ChartSeries series)
        {
            if (!series.HasBarBorder || !series.HasVisibleLine || bounds.IsEmpty) return;
            using (var paint = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = series.LineColor,
                StrokeWidth = (float)series.LineWidthPt,
                IsAntialias = true,
                StrokeJoin = SKStrokeJoin.Miter
            })
            {
                canvas.DrawRect(bounds, paint);
            }
        }
    }
}
