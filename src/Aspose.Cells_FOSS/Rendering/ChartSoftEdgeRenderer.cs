using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartSoftEdgeRenderer
    {
        public static void DrawSolidFill(SKCanvas canvas, SKRect area, SKColor color, double radiusPt)
        {
            var radius = (float)Math.Min(radiusPt, Math.Min(area.Width, area.Height) * 0.4d);
            var inset = radius * 0.8f;
            var inner = new SKRect(area.Left + inset, area.Top + inset,
                area.Right - inset, area.Bottom - inset);
            using (var blur = SKMaskFilter.CreateBlur(SKBlurStyle.Normal, radius / 3f))
            using (var paint = new SKPaint { Style = SKPaintStyle.Fill, Color = color,
                IsAntialias = true, MaskFilter = blur })
            {
                canvas.Save();
                canvas.ClipRect(area);
                canvas.DrawRect(inner, paint);
                canvas.Restore();
            }
        }
    }
}
