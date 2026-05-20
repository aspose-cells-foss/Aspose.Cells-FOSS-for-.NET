using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents page setup model.
    /// </summary>
    internal sealed class PageSetupModel
    {
        /// <summary>
        /// Gets the page margin settings.
        /// </summary>
        public PageMarginsModel Margins { get; } = new PageMarginsModel();
        /// <summary>
        /// Gets the print option settings.
        /// </summary>
        public PrintOptionsModel PrintOptions { get; } = new PrintOptionsModel();
        /// <summary>
        /// Gets the header and footer settings.
        /// </summary>
        public HeaderFooterModel HeaderFooter { get; } = new HeaderFooterModel();
        /// <summary>
        /// Gets or sets the paper size.
        /// </summary>
        public int PaperSize { get; set; }
        /// <summary>
        /// Gets or sets the orientation.
        /// </summary>
        public PageOrientation Orientation { get; set; }
        /// <summary>
        /// Gets or sets the first page number.
        /// </summary>
        public int? FirstPageNumber { get; set; }
        /// <summary>
        /// Gets or sets the scale.
        /// </summary>
        public int? Scale { get; set; }
        /// <summary>
        /// Gets or sets the fit to width.
        /// </summary>
        public int? FitToWidth { get; set; }
        /// <summary>
        /// Gets or sets the fit to height.
        /// </summary>
        public int? FitToHeight { get; set; }
        /// <summary>
        /// Gets or sets the print area.
        /// </summary>
        public string PrintArea { get; set; }
        /// <summary>
        /// Gets or sets the print title rows.
        /// </summary>
        public string PrintTitleRows { get; set; }
        /// <summary>
        /// Gets or sets the print title columns.
        /// </summary>
        public string PrintTitleColumns { get; set; }
        /// <summary>
        /// Gets the horizontal manual page break row indexes.
        /// </summary>
        public List<int> HorizontalPageBreaks { get; } = new List<int>();
        /// <summary>
        /// Gets the vertical manual page break column indexes.
        /// </summary>
        public List<int> VerticalPageBreaks { get; } = new List<int>();
    }
}

