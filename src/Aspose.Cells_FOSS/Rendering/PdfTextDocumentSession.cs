using System.Collections.Generic;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfTextDocumentSession
    {
        private readonly List<PdfTextRunRecord> _runs = new List<PdfTextRunRecord>();

        public bool HasRuns
        {
            get { return _runs.Count > 0; }
        }

        public IList<PdfTextRunRecord> Runs
        {
            get { return _runs; }
        }

        public void RecordWorksheetRun(PageLayout page, SKRect clipRect, string text, SKTypeface typeface, SKColor color, float fontSizePt, float xPt, float baselinePt)
        {
            RecordPageRun(page, clipRect, text, typeface, color, fontSizePt, xPt, baselinePt);
        }

        public void RecordChartRun(PageLayout page, SKRect clipRect, string text, SKTypeface typeface, SKColor color, float fontSizePt, float xPt, float baselinePt)
        {
            if (page == null || typeface == null || string.IsNullOrEmpty(text))
            {
                return;
            }

            var scale = (float)page.ScaleFactor;
            var x = (float)page.ContentOriginXPt + scale * xPt;
            var baselineTop = (float)page.ContentOriginYPt + scale * baselinePt;
            var clipLeft = (float)page.ContentOriginXPt + scale * clipRect.Left;
            var clipTop = (float)page.ContentOriginYPt + scale * clipRect.Top;
            var clipRight = (float)page.ContentOriginXPt + scale * clipRect.Right;
            var clipBottom = (float)page.ContentOriginYPt + scale * clipRect.Bottom;

            _runs.Add(new PdfTextRunRecord
            {
                PageNumber = page.PageNumber,
                Text = text,
                Typeface = typeface,
                // Chart text replays inside an explicit top-down page transform, so its Type3 glyphs
                // need the positive-Y font matrix to stay upright in that coordinate space.
                UsePositiveYFontMatrix = true,
                UseTopDownCoordinates = true,
                Color = color,
                FontSizePt = scale * fontSizePt,
                XPt = x,
                BaselineYPt = baselineTop,
                PageHeightPt = (float)page.PageHeightPt,
                ClipLeftPt = clipLeft,
                ClipTopPt = clipTop,
                ClipBottomPt = clipBottom,
                ClipWidthPt = clipRight - clipLeft,
                ClipHeightPt = clipBottom - clipTop,
            });
        }

        public void RecordChartRotatedRun(PageLayout page, SKRect clipRect, string text, SKTypeface typeface, SKColor color, float fontSizePt, float originXPt, float originYPt, float localXPt, float localBaselinePt, float rotationDeg)
        {
            if (page == null || typeface == null || string.IsNullOrEmpty(text))
            {
                return;
            }

            var scale = (float)page.ScaleFactor;
            var originX = (float)page.ContentOriginXPt + scale * originXPt;
            var originY = (float)page.ContentOriginYPt + scale * originYPt;
            var clipLeft = (float)page.ContentOriginXPt + scale * clipRect.Left;
            var clipTop = (float)page.ContentOriginYPt + scale * clipRect.Top;
            var clipRight = (float)page.ContentOriginXPt + scale * clipRect.Right;
            var clipBottom = (float)page.ContentOriginYPt + scale * clipRect.Bottom;

            _runs.Add(new PdfTextRunRecord
            {
                PageNumber = page.PageNumber,
                Text = text,
                Typeface = typeface,
                UsePositiveYFontMatrix = true,
                UseTopDownCoordinates = true,
                Color = color,
                FontSizePt = scale * fontSizePt,
                XPt = scale * localXPt,
                BaselineYPt = scale * localBaselinePt,
                PageHeightPt = (float)page.PageHeightPt,
                ClipLeftPt = clipLeft,
                ClipTopPt = clipTop,
                ClipBottomPt = clipBottom,
                ClipWidthPt = clipRight - clipLeft,
                ClipHeightPt = clipBottom - clipTop,
                TransformOriginXPt = originX,
                TransformOriginYPt = originY,
                RotationDeg = rotationDeg,
            });
        }

        public void RecordPageRun(PageLayout page, SKRect clipRect, string text, SKTypeface typeface, SKColor color, float fontSizePt, float xPt, float baselinePt)
        {
            if (page == null || typeface == null || string.IsNullOrEmpty(text))
            {
                return;
            }

            var scale = (float)page.ScaleFactor;
            var x = (float)page.ContentOriginXPt + scale * xPt;
            var pageHeight = (float)page.PageHeightPt;
            var baselineTopDown = (float)page.ContentOriginYPt + scale * baselinePt;
            var baseline = pageHeight - baselineTopDown;

            var clipLeft = (float)page.ContentOriginXPt + scale * clipRect.Left;
            var clipTop = (float)page.ContentOriginYPt + scale * clipRect.Top;
            var clipRight = (float)page.ContentOriginXPt + scale * clipRect.Right;
            var clipBottom = (float)page.ContentOriginYPt + scale * clipRect.Bottom;
            var clipBottomPdf = pageHeight - clipBottom;

            _runs.Add(new PdfTextRunRecord
            {
                PageNumber = page.PageNumber,
                Text = text,
                Typeface = typeface,
                Color = color,
                FontSizePt = scale * fontSizePt,
                XPt = x,
                BaselineYPt = baseline,
                PageHeightPt = pageHeight,
                ClipLeftPt = clipLeft,
                ClipTopPt = clipTop,
                ClipBottomPt = clipBottomPdf,
                ClipWidthPt = clipRight - clipLeft,
                ClipHeightPt = clipBottom - clipTop,
            });
        }
    }
}
