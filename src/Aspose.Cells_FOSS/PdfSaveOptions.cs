using System;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Options controlling XLSX-to-PDF rendering. Mirrors the most common settings of
    /// Aspose.Cells' <c>PdfSaveOptions</c>. Page geometry (paper size, orientation, margins) is not
    /// set here — configure it through <see cref="Worksheet.PageSetup"/>, matching Aspose.Cells.
    /// </summary>
    /// <example>
    /// <code>
    /// // Fit every worksheet onto a single page.
    /// workbook.Save("report.pdf", new PdfSaveOptions { OnePagePerSheet = true });
    ///
    /// // Paper size / orientation are configured via the worksheet's page setup.
    /// workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperA4;
    /// workbook.Save("report.pdf", SaveFormat.Pdf);
    /// </code>
    /// </example>
    public sealed class PdfSaveOptions : SaveOptions
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PdfSaveOptions"/> class with
        /// <see cref="SaveOptions.SaveFormat"/> set to <see cref="SaveFormat.Pdf"/>.
        /// </summary>
        public PdfSaveOptions()
        {
            SaveFormat = SaveFormat.Pdf;
            UseType3TextOptimization = true;
            UseTrueTypeSubsetPrototype = true;
        }

        /// <summary>
        /// Gets or sets whether all content of each worksheet is rendered onto a single PDF page.
        /// When <see langword="true"/>, the sheet is scaled to fit both the page width and height and
        /// is never split across pages.
        /// </summary>
        public bool OnePagePerSheet { get; set; }

        /// <summary>
        /// Gets or sets whether all columns of each worksheet are rendered on a single page width.
        /// When <see langword="true"/>, the sheet is scaled so every column fits horizontally on one
        /// page while rows may still flow onto additional pages. Ignored when
        /// <see cref="OnePagePerSheet"/> is <see langword="true"/>.
        /// </summary>
        public bool AllColumnsInOnePagePerSheet { get; set; }

        /// <summary>
        /// Gets or sets the fallback font family used when a cell's configured font cannot be
        /// resolved on the host system. When null or empty, the renderer's built-in fallback chain
        /// is used.
        /// </summary>
        public string DefaultFont { get; set; }

        internal bool UseType3TextOptimization { get; set; }
        internal bool UseTrueTypeSubsetPrototype { get; set; }
    }
}
