using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents page setup model.
    /// </summary>
    public sealed class PageSetupModel
    {
        /// <summary>
        /// Performs page margins model.
        /// </summary>
        /// <returns>The page margins model margins { get; } = new.</returns>
        public PageMarginsModel Margins { get; } = new PageMarginsModel();
        /// <summary>
        /// Performs print options model.
        /// </summary>
        /// <returns>The print options model print options { get; } = new.</returns>
        public PrintOptionsModel PrintOptions { get; } = new PrintOptionsModel();
        /// <summary>
        /// Performs header footer model.
        /// </summary>
        /// <returns>The header footer model header footer { get; } = new.</returns>
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
        /// Performs list<int>.
        /// </summary>
        /// <returns>The list<int> horizontal page breaks { get; } = new.</returns>
        public List<int> HorizontalPageBreaks { get; } = new List<int>();
        /// <summary>
        /// Performs list<int>.
        /// </summary>
        /// <returns>The list<int> vertical page breaks { get; } = new.</returns>
        public List<int> VerticalPageBreaks { get; } = new List<int>();
    }
}
