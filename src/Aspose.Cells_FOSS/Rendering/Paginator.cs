using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Turns a sheet's resolved <see cref="SheetLayout"/> plus its page setup into a flat list of
    /// <see cref="PageLayout"/> pages, applying scale/fit-to-page and greedy row/column splitting.
    /// </summary>
    internal static class Paginator
    {
        public static List<PageLayout> Paginate(SheetLayout layout, int sheetIndex, PdfSaveOptions options)
        {
            var pages = new List<PageLayout>();
            if (layout.IsEmpty)
            {
                return pages;
            }

            if (options == null)
            {
                options = new PdfSaveOptions();
            }

            var setup = layout.Sheet.PageSetup;

            double pageWidthPt, pageHeightPt;
            PaperSizes.GetDimensionsPoints(setup.PaperSize, out pageWidthPt, out pageHeightPt);
            if (setup.Orientation == PageOrientation.Landscape)
            {
                var swap = pageWidthPt; pageWidthPt = pageHeightPt; pageHeightPt = swap;
            }

            var leftPt = RenderUnits.InchesToPoints(setup.Margins.Left);
            var rightPt = RenderUnits.InchesToPoints(setup.Margins.Right);
            var topPt = RenderUnits.InchesToPoints(setup.Margins.Top);
            var bottomPt = RenderUnits.InchesToPoints(setup.Margins.Bottom);

            var usableWidthPt = Math.Max(1d, pageWidthPt - leftPt - rightPt);
            var usableHeightPt = Math.Max(1d, pageHeightPt - topPt - bottomPt);

            var totalWidthPt = layout.ColumnStartPt[layout.LastColumn + 1] - layout.ColumnStartPt[layout.FirstColumn];
            var totalHeightPt = layout.RowStartPt[layout.LastRow + 1] - layout.RowStartPt[layout.FirstRow];

            var forceOnePageWide = options.OnePagePerSheet || options.AllColumnsInOnePagePerSheet;
            var forceOnePageTall = options.OnePagePerSheet;

            var scale = ResolveScale(setup, options, totalWidthPt, totalHeightPt, usableWidthPt, usableHeightPt);

            var colBreaks = forceOnePageWide
                ? SingleSpan(layout.FirstColumn, layout.LastColumn)
                : SplitAxis(
                    layout.FirstColumn,
                    layout.LastColumn,
                    usableWidthPt / scale,
                    delegate(int index) { return layout.ColumnWidthPt[index]; },
                    setup.VerticalPageBreaks);

            var rowBreaks = forceOnePageTall
                ? SingleSpan(layout.FirstRow, layout.LastRow)
                : SplitAxis(
                    layout.FirstRow,
                    layout.LastRow,
                    usableHeightPt / scale,
                    delegate(int index) { return layout.RowHeightPt[index]; },
                    setup.HorizontalPageBreaks);

            foreach (var rowSpan in rowBreaks)
            {
                foreach (var colSpan in colBreaks)
                {
                    var page = new PageLayout();
                    page.SheetIndex = sheetIndex;
                    page.Sheet = layout;
                    page.StartRow = rowSpan.Start;
                    page.EndRow = rowSpan.End;
                    page.StartColumn = colSpan.Start;
                    page.EndColumn = colSpan.End;
                    page.ScaleFactor = scale;
                    page.PageWidthPt = pageWidthPt;
                    page.PageHeightPt = pageHeightPt;
                    page.ContentOriginXPt = leftPt;
                    page.ContentOriginYPt = topPt;
                    pages.Add(page);
                }
            }

            return pages;
        }

        private static double ResolveScale(PageSetupModel setup, PdfSaveOptions options, double totalWidthPt, double totalHeightPt, double usableWidthPt, double usableHeightPt)
        {
            // Save-option fit modes take precedence over the sheet's own page-setup scaling.
            if (options.OnePagePerSheet)
            {
                var w = totalWidthPt > 0d ? usableWidthPt / totalWidthPt : 1d;
                var h = totalHeightPt > 0d ? usableHeightPt / totalHeightPt : 1d;
                return Math.Min(1d, Math.Min(w, h));
            }

            if (options.AllColumnsInOnePagePerSheet)
            {
                var w = totalWidthPt > 0d ? usableWidthPt / totalWidthPt : 1d;
                return Math.Min(1d, w);
            }

            var fitWidth = setup.FitToWidth.HasValue && setup.FitToWidth.Value > 0;
            var fitHeight = setup.FitToHeight.HasValue && setup.FitToHeight.Value > 0;

            if (fitWidth || fitHeight)
            {
                var scale = double.MaxValue;
                if (fitWidth && totalWidthPt > 0d)
                {
                    scale = Math.Min(scale, setup.FitToWidth.Value * usableWidthPt / totalWidthPt);
                }

                if (fitHeight && totalHeightPt > 0d)
                {
                    scale = Math.Min(scale, setup.FitToHeight.Value * usableHeightPt / totalHeightPt);
                }

                if (scale == double.MaxValue)
                {
                    return 1d;
                }

                // Fit-to-page only ever shrinks content; it never enlarges past 100%.
                return Math.Min(1d, scale);
            }

            if (setup.Scale.HasValue && setup.Scale.Value > 0)
            {
                return setup.Scale.Value / 100d;
            }

            return 1d;
        }

        private static List<Span> SingleSpan(int first, int last)
        {
            return new List<Span> { new Span(first, last) };
        }

        private delegate double SizeAccessor(int index);

        private struct Span
        {
            public int Start;
            public int End;
            public Span(int start, int end) { Start = start; End = end; }
        }

        /// <summary>
        /// Greedily groups indices 0..last into spans that each fit within <paramref name="capacity"/>
        /// (in unscaled points), honoring manual breaks that fall earlier than the greedy break.
        /// </summary>
        private static List<Span> SplitAxis(int first, int last, double capacity, SizeAccessor size, List<int> manualBreaks)
        {
            var breakSet = new HashSet<int>();
            if (manualBreaks != null)
            {
                foreach (var b in manualBreaks)
                {
                    breakSet.Add(b);
                }
            }
            var spans = new List<Span>();
            var start = first;
            var accumulated = 0d;

            for (var index = first; index <= last; index++)
            {
                var cellSize = size(index);

                var manualBreakHere = index > start && breakSet.Contains(index);
                var overflow = index > start && accumulated + cellSize > capacity;

                if (manualBreakHere || overflow)
                {
                    spans.Add(new Span(start, index - 1));
                    start = index;
                    accumulated = 0d;
                }

                accumulated += cellSize;
            }

            spans.Add(new Span(start, last));
            return spans;
        }
    }
}
