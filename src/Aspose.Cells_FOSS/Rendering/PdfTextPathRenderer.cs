using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Draws text as vector outlines so PDF output avoids embedding full font files.
    /// </summary>
    internal static class PdfTextPathRenderer
    {
        public static void DrawText(SKCanvas canvas, string text, float x, float baseline, SKFont font, SKColor color)
        {
            if (canvas == null || font == null || string.IsNullOrEmpty(text))
            {
                return;
            }

            using (var paint = new SKPaint(font))
            {
                paint.Color = color;
                paint.Style = SKPaintStyle.Fill;
                paint.IsAntialias = true;

                using (var path = paint.GetTextPath(text, x, baseline))
                {
                    if (path != null)
                    {
                        canvas.DrawPath(path, paint);
                    }
                }
            }
        }
    }
}
