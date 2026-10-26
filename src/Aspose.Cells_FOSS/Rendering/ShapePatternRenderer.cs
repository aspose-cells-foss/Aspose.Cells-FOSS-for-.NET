using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapePatternRenderer
    {
        public static void Draw(SKCanvas canvas, SKRect rect, SKPath outline, ParsedShape shape)
        {
            if (shape.FillPattern != "pct5") return;
            // DrawingML pct5: two 2x2 dots in a 16x16 tile, rendered at 192 dpi.
            const float tile = 6f;
            const float dot = 0.75f;
            canvas.Save();
            canvas.ClipPath(outline, SKClipOperation.Intersect, true);
            using (var paint = new SKPaint { Color = shape.PatternForeground, IsAntialias = true })
            {
                for (var y = rect.Top; y < rect.Bottom; y += tile)
                {
                    for (var x = rect.Left; x < rect.Right; x += tile)
                    {
                        canvas.DrawRect(x, y, dot, dot, paint);
                        canvas.DrawRect(x + tile / 2f, y + tile / 2f, dot, dot, paint);
                    }
                }
            }
            canvas.Restore();
        }
    }
}
