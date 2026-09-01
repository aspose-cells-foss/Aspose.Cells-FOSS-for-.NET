using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class RichTextCellRenderer
    {
        private readonly RenderContext _context;

        public RichTextCellRenderer(RenderContext context)
        {
            _context = context;
        }

        public bool CanRender(CellRecord record, StyleValue style, string text)
        {
            if (record == null || record.RichTextRuns == null || record.RichTextRuns.Count == 0)
            {
                return false;
            }

            if (style == null)
            {
                return false;
            }

            if (style.Alignment.WrapText)
            {
                return false;
            }

            if (style.Alignment.TextRotation != 0)
            {
                return false;
            }

            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            return text.IndexOf('\r') < 0 && text.IndexOf('\n') < 0;
        }

        public float MeasureWidth(CellRecord record, StyleValue style, string text, bool useOverrideColor, SKColor overrideColor)
        {
            var slices = BuildSlices(record, style, text, useOverrideColor, overrideColor);
            var total = 0f;
            for (var i = 0; i < slices.Count; i++)
            {
                total += slices[i].WidthPt;
            }

            return total;
        }

        public bool TryDraw(SKCanvas canvas, PageLayout page, SKRect rect, SKRect clip, CellRecord record, StyleValue style, HorizontalAlignment horizontal, bool useOverrideColor, SKColor overrideColor)
        {
            var text = record.Value as string;
            if (!CanRender(record, style, text))
            {
                return false;
            }

            var slices = BuildSlices(record, style, text, useOverrideColor, overrideColor);
            if (slices.Count == 0)
            {
                return false;
            }

            var totalWidth = 0f;
            var maxAscent = 0f;
            var maxDescent = 0f;
            for (var i = 0; i < slices.Count; i++)
            {
                var slice = slices[i];
                totalWidth += slice.WidthPt;
                if (slice.AscentPt > maxAscent)
                {
                    maxAscent = slice.AscentPt;
                }

                if (slice.DescentPt > maxDescent)
                {
                    maxDescent = slice.DescentPt;
                }
            }

            var lineHeight = maxAscent + maxDescent;
            if (lineHeight <= 0f)
            {
                lineHeight = _context.GetFontContext(style.Font).LineHeightPt;
            }

            float blockTop;
            switch (style.Alignment.Vertical)
            {
                case VerticalAlignment.Top:
                    blockTop = rect.Top + (float)SheetLayout.VerticalPaddingPt;
                    break;
                case VerticalAlignment.Center:
                    blockTop = rect.Top + (rect.Height - lineHeight) / 2f;
                    break;
                default:
                    blockTop = rect.Bottom - (float)SheetLayout.VerticalPaddingPt - lineHeight;
                    break;
            }

            var padding = (float)SheetLayout.HorizontalPaddingPt;
            var indentOffset = IndentOffsetPt(_context.GetFontContext(style.Font), style.Alignment, horizontal);
            float x;
            switch (horizontal)
            {
                case HorizontalAlignment.Center:
                case HorizontalAlignment.CenterContinuous:
                    x = rect.Left + (rect.Width - totalWidth) / 2f;
                    break;
                case HorizontalAlignment.Right:
                    x = rect.Right - padding - indentOffset - totalWidth;
                    break;
                default:
                    x = rect.Left + padding + indentOffset;
                    break;
            }

            var baseline = blockTop + maxAscent;
            canvas.Save();
            canvas.ClipRect(clip);

            using (var paint = new SKPaint())
            {
                paint.IsAntialias = true;

                for (var i = 0; i < slices.Count; i++)
                {
                    var slice = slices[i];
                    var runs = slice.FontContext.SplitRuns(slice.Text);
                    var cursor = x;
                    paint.Color = slice.Color;

                    for (var runIndex = 0; runIndex < runs.Count; runIndex++)
                    {
                        var run = runs[runIndex];
                        using (var font = new SKFont(run.Typeface, slice.FontContext.SizePt))
                        {
                            font.Subpixel = true;
                            PdfTextPathRenderer.DrawText(canvas, run.Text, cursor, baseline, font, paint.Color);
                        }

                        DrawDecorations(canvas, paint, cursor, baseline, run.WidthPt, slice.FontContext.SizePt, slice.Font);
                        cursor += run.WidthPt;
                    }

                    x += slice.WidthPt;
                }
            }

            canvas.Restore();
            return true;
        }

        private static float IndentOffsetPt(CellFontContext fontContext, AlignmentValue alignment, HorizontalAlignment horizontal)
        {
            if (fontContext == null || alignment == null)
            {
                return 0f;
            }

            if (alignment.IndentLevel <= 0)
            {
                return 0f;
            }

            if (horizontal != HorizontalAlignment.Left
                && horizontal != HorizontalAlignment.Right
                && horizontal != HorizontalAlignment.Distributed
                && horizontal != HorizontalAlignment.Justify)
            {
                return 0f;
            }

            var digitWidth = fontContext.Measure("0");
            var indentStep = Math.Max(fontContext.SizePt * 0.68f, digitWidth * 1.15f);
            return alignment.IndentLevel * indentStep;
        }

        private List<RichTextSlice> BuildSlices(CellRecord record, StyleValue style, string text, bool useOverrideColor, SKColor overrideColor)
        {
            var slices = new List<RichTextSlice>();
            var currentIndex = 0;
            var runs = record.RichTextRuns;
            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                if (run.StartIndex > currentIndex)
                {
                    AddSlice(slices, text.Substring(currentIndex, run.StartIndex - currentIndex), style.Font, useOverrideColor, overrideColor);
                }

                var runFont = MergeFont(style.Font, run.Font);
                AddSlice(slices, text.Substring(run.StartIndex, run.Length), runFont, useOverrideColor, overrideColor);
                currentIndex = run.StartIndex + run.Length;
            }

            if (currentIndex < text.Length)
            {
                AddSlice(slices, text.Substring(currentIndex), style.Font, useOverrideColor, overrideColor);
            }

            return slices;
        }

        private void AddSlice(List<RichTextSlice> slices, string text, FontValue font, bool useOverrideColor, SKColor overrideColor)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var resolvedFont = font ?? new FontValue();
            var context = _context.GetFontContext(resolvedFont);
            var metrics = MeasureMetrics(resolvedFont, context);
            slices.Add(new RichTextSlice
            {
                Text = text,
                Font = resolvedFont,
                FontContext = context,
                WidthPt = context.Measure(text),
                AscentPt = metrics.Item1,
                DescentPt = metrics.Item2,
                Color = useOverrideColor ? overrideColor : _context.Colors.Resolve(resolvedFont.Color, SKColors.Black),
            });
        }

        private Tuple<float, float> MeasureMetrics(FontValue font, CellFontContext context)
        {
            using (var skFont = new SKFont(_context.Fonts.Resolve(font), context.SizePt))
            {
                var metrics = skFont.Metrics;
                return Tuple.Create(-metrics.Ascent, metrics.Descent);
            }
        }

        private FontValue MergeFont(FontValue baseFont, FontValue richFont)
        {
            if (baseFont == null && richFont == null)
            {
                return new FontValue();
            }

            if (baseFont == null)
            {
                return richFont.Clone();
            }

            if (richFont == null)
            {
                return baseFont.Clone();
            }

            var merged = baseFont.Clone();
            merged.Name = richFont.Name;
            merged.Size = richFont.Size;
            merged.Bold = richFont.Bold;
            merged.Italic = richFont.Italic;
            merged.Underline = richFont.Underline;
            merged.StrikeThrough = richFont.StrikeThrough;
            merged.Color = richFont.Color;
            merged.Family = richFont.Family;
            merged.Scheme = richFont.Scheme;
            return merged;
        }

        private void DrawDecorations(SKCanvas canvas, SKPaint paint, float x, float baseline, float width, float fontSize, FontValue font)
        {
            if (font != null && font.Underline != FontUnderlineType.None)
            {
                using (var linePaint = new SKPaint())
                {
                    linePaint.Color = paint.Color;
                    linePaint.Style = SKPaintStyle.Stroke;
                    linePaint.IsAntialias = true;
                    linePaint.StrokeWidth = Math.Max(0.6f, fontSize * 0.055f);
                    var underlineY = baseline + fontSize * 0.08f;
                    canvas.DrawLine(x, underlineY, x + width, underlineY, linePaint);
                }
            }

            if (font != null && font.StrikeThrough)
            {
                using (var linePaint = new SKPaint())
                {
                    linePaint.Color = paint.Color;
                    linePaint.Style = SKPaintStyle.Stroke;
                    linePaint.IsAntialias = true;
                    linePaint.StrokeWidth = Math.Max(0.55f, fontSize * 0.05f);
                    var strikeY = baseline - fontSize * 0.28f;
                    canvas.DrawLine(x, strikeY, x + width, strikeY, linePaint);
                }
            }
        }
    }
}
