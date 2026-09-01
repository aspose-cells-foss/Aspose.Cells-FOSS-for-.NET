namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// A single physical PDF page: which slice of a sheet it shows, the scale to apply, and the
    /// page/margin geometry in points. Pagination fully resolves these before rendering starts, so
    /// the renderer never needs to know about sibling pages or scale-solving.
    /// </summary>
    internal sealed class PageLayout
    {
        public int SheetIndex;
        public SheetLayout Sheet;

        public int StartRow;
        public int EndRow;   // inclusive
        public int StartColumn;
        public int EndColumn; // inclusive

        public double ScaleFactor = 1d;

        public double PageWidthPt;
        public double PageHeightPt;

        // Top-left of the printable content area (i.e. the margins), in points.
        public double ContentOriginXPt;
        public double ContentOriginYPt;

        // 1-based page index and total, for header/footer substitution.
        public int PageNumber;
        public int TotalPages;
    }
}
