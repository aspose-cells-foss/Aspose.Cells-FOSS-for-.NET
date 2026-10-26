using System;
using System.Collections.Generic;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapeWrappedTextRenderer
    {
        public static void Draw(SKCanvas canvas, SKRect rect, ParsedShape shape, CellFontContext fontContext, SKPaint paint)
        {
            var availableWidth = Math.Max(1f, rect.Width - (float)shape.TextInsetLeftPt * 2f);
            var lines = new List<string>();
            var words = shape.Text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var current = string.Empty;
            foreach (var word in words)
            {
                var candidate = current.Length == 0 ? word : current + " " + word;
                if (current.Length > 0 && fontContext.Measure(candidate) > availableWidth)
                {
                    lines.Add(current);
                    current = word;
                }
                else
                {
                    current = candidate;
                }
            }

            if (current.Length > 0)
            {
                lines.Add(current);
            }

            if (lines.Count == 0)
            {
                return;
            }

            var metrics = paint.FontMetrics;
            var lineHeight = paint.TextSize * 1.2f;
            var textHeight = (lines.Count - 1) * lineHeight + metrics.Descent - metrics.Ascent;
            float baseline;
            if (shape.TextVAnchor == "t") baseline = rect.Top + (float)shape.TextInsetTopPt - metrics.Ascent;
            else if (shape.TextVAnchor == "b") baseline = rect.Bottom - (float)shape.TextInsetTopPt - textHeight - metrics.Ascent;
            else baseline = rect.MidY - textHeight / 2f - metrics.Ascent;

            canvas.Save();
            canvas.ClipRect(rect);
            using (var fontObject = new SKFont(paint.Typeface, paint.TextSize))
            {
                fontObject.Subpixel = true;
                foreach (var line in lines)
                {
                    var lineWidth = fontContext.Measure(line);
                    float x;
                    if (shape.TextHAlign == "l") x = rect.Left + (float)shape.TextInsetLeftPt;
                    else if (shape.TextHAlign == "r") x = rect.Right - (float)shape.TextInsetLeftPt - lineWidth;
                    else x = rect.MidX - lineWidth / 2f;
                    PdfTextPathRenderer.DrawText(canvas, line, x, baseline, fontObject, paint.Color);
                    baseline += lineHeight;
                }
            }
            canvas.Restore();
        }
    }
}
