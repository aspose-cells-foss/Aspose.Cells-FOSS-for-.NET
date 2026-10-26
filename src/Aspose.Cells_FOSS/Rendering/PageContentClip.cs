using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class PageContentClip
    {
        public static SKRect Resolve(PageLayout page, bool floatingObjects)
        {
            var layout = page.Sheet;
            var width = (float)(layout.ColumnStartPt[page.EndColumn + 1]
                - layout.ColumnStartPt[page.StartColumn]);
            var height = (float)(layout.RowStartPt[page.EndRow + 1]
                - layout.RowStartPt[page.StartRow]);
            // Floating objects may overrun the used range, but never a page split.
            if (floatingObjects)
            {
                if (page.EndColumn >= layout.LastColumn) width += 72f;
                if (page.EndRow >= layout.LastRow) height += 72f;
            }
            return new SKRect(0f, 0f, width, height);
        }

        public static SKRect Intersect(PageLayout page, SKRect clip, bool floatingObjects)
        {
            if (page.Sheet == null)
            {
                return clip;
            }
            return SKRect.Intersect(clip, Resolve(page, floatingObjects));
        }
    }
}
