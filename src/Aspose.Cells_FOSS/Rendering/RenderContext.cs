using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Bundles the services shared across the layout and rendering passes for a single export, and
    /// caches font contexts keyed by their visual configuration so both passes measure and draw text
    /// with identical metrics.
    /// </summary>
    internal sealed class RenderContext : IDisposable
    {
        private readonly Dictionary<string, CellFontContext> _fontContexts = new Dictionary<string, CellFontContext>(StringComparer.Ordinal);

        public RenderContext(WorkbookModel workbook)
            : this(workbook, null)
        {
        }

        public RenderContext(WorkbookModel workbook, string defaultFont)
        {
            Workbook = workbook;
            Fonts = new FontRegistry();
            Pictures = new PicturePdfImageCache();
            Charts = new ChartPdfImageCache();
            Fonts.DefaultFontName = defaultFont;
            Colors = RenderColor.FromWorkbook(workbook);
            // Match Excel, which renders locale-dependent formats (e.g. the short-date built-in) using
            // the viewer's regional settings: use the machine culture unless the workbook pins an
            // explicit (non-invariant) display culture.
            var displayCulture = workbook != null && workbook.Settings != null ? workbook.Settings.DisplayCulture : null;
            Culture = displayCulture != null && !CultureInfo.InvariantCulture.Equals(displayCulture)
                ? displayCulture
                : CultureInfo.CurrentCulture;
            var normalFont = workbook != null && workbook.DefaultStyle != null ? workbook.DefaultStyle.Font : null;
            MaxDigitWidthPixels = Fonts.MeasureMaxDigitWidth(normalFont);
            FontDerivedRowHeightPt = Fonts.MeasureDefaultRowHeightPt(normalFont);
        }

        public WorkbookModel Workbook { get; private set; }
        public FontRegistry Fonts { get; private set; }
        public PicturePdfImageCache Pictures { get; private set; }
        public ChartPdfImageCache Charts { get; private set; }
        public RenderColor Colors { get; private set; }
        public CultureInfo Culture { get; private set; }
        public double MaxDigitWidthPixels { get; private set; }
        public double FontDerivedRowHeightPt { get; private set; }
        public PdfTextDocumentSession PdfTextSession { get; set; }
        public bool EnableWorksheetTextOptimization { get; set; }

        public CellFontContext GetFontContext(FontValue font)
        {
            var name = font != null && font.Name != null ? font.Name : "Calibri";
            var size = font != null ? font.Size : 11d;
            var bold = font != null && font.Bold;
            var italic = font != null && font.Italic;
            var key = name + "|" + size.ToString("R", CultureInfo.InvariantCulture) + "|" + (bold ? "1" : "0") + "|" + (italic ? "1" : "0");

            CellFontContext context;
            if (_fontContexts.TryGetValue(key, out context))
            {
                return context;
            }

            context = new CellFontContext(Fonts, font);
            _fontContexts[key] = context;
            return context;
        }

        public void Dispose()
        {
            foreach (var context in _fontContexts.Values)
            {
                context.Dispose();
            }

            _fontContexts.Clear();
            Charts.Dispose();
            Pictures.Dispose();
            Fonts.Dispose();
        }
    }
}
