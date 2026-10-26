using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ShapeVerticalTextRenderer
    {
        public static bool IsVerticalText(ParsedShape shape)
        {
            if (shape == null || string.IsNullOrEmpty(shape.TextVerticalType))
            {
                return false;
            }

            return shape.TextVerticalType == "vert"
                || shape.TextVerticalType == "vert270"
                || shape.TextVerticalType == "wordArtVert"
                || shape.TextVerticalType == "wordArtVertRtl";
        }

        public static void Draw(SKCanvas canvas, SKRect rect, ParsedShape shape, CellFontContext fontContext, SKPaint paint)
        {
            if (shape.TextVerticalType == "wordArtVert" || shape.TextVerticalType == "wordArtVertRtl")
            {
                DrawWordArtVerticalText(canvas, rect, shape, fontContext, paint);
                return;
            }

            var clockwise = shape.TextVerticalType == "vert270" || shape.TextVerticalType == "wordArtVertRtl";
            canvas.Save();
            canvas.Translate(rect.MidX, rect.MidY);
            canvas.RotateDegrees(clockwise ? 90f : -90f);

            var metrics = paint.FontMetrics;
            var width = fontContext.Measure(shape.Text);
            var drawX = -width / 2f;
            var baseline = -(metrics.Ascent + metrics.Descent) / 2f;
            using (var fontObject = new SKFont(paint.Typeface, paint.TextSize))
            {
                fontObject.Subpixel = true;
                PdfTextPathRenderer.DrawText(canvas, shape.Text, drawX, baseline, fontObject, paint.Color);
            }

            canvas.Restore();
        }

        private static void DrawWordArtVerticalText(SKCanvas canvas, SKRect rect, ParsedShape shape, CellFontContext fontContext, SKPaint paint)
        {
            if (string.IsNullOrEmpty(shape.Text))
            {
                return;
            }

            var metrics = paint.FontMetrics;
            var charAdvance = Math.Max((float)shape.TextSizePt, metrics.Descent - metrics.Ascent);
            var availableHeight = Math.Max(charAdvance, rect.Height - 2f * (float)shape.TextInsetTopPt);
            var rows = Math.Max(1, (int)Math.Floor(availableHeight / charAdvance));
            var columns = (shape.Text.Length + rows - 1) / rows;
            var totalHeight = charAdvance * Math.Min(rows, shape.Text.Length);
            var currentY = rect.MidY - totalHeight * 0.5f;

            canvas.Save();
            canvas.ClipRect(rect);
            using (var fontObject = new SKFont(paint.Typeface, paint.TextSize))
            {
                fontObject.Subpixel = true;
                for (var i = 0; i < shape.Text.Length; i++)
                {
                    var ch = shape.Text[i].ToString();
                    var column = i / rows;
                    if (shape.TextVerticalType == "wordArtVertRtl") column = columns - 1 - column;
                    var centerX = rect.MidX + (column - (columns - 1) * 0.5f) * charAdvance;
                    var width = fontContext.Measure(ch);
                    var drawX = centerX - width * 0.5f;
                    var baseline = currentY + (i % rows) * charAdvance - metrics.Ascent;
                    PdfTextPathRenderer.DrawText(canvas, ch, drawX, baseline, fontObject, paint.Color);
                }
            }
            canvas.Restore();
        }
    }
}
