using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents worksheet view model.
    /// </summary>
    internal sealed class WorksheetViewModel
    {
        /// <summary>
        /// Gets or sets a value indicating whether show grid lines.
        /// </summary>
        public bool ShowGridLines { get; set; } = true;
        /// <summary>
        /// Gets or sets a value indicating whether show row column headers.
        /// </summary>
        public bool ShowRowColumnHeaders { get; set; } = true;
        /// <summary>
        /// Gets or sets a value indicating whether show zeros.
        /// </summary>
        public bool ShowZeros { get; set; } = true;
        /// <summary>
        /// Gets or sets a value indicating whether right to left.
        /// </summary>
        public bool RightToLeft { get; set; }
        /// <summary>
        /// Gets or sets the zoom scale.
        /// </summary>
        public int ZoomScale { get; set; } = 100;
        /// <summary>
        /// Gets or sets the sheet view mode such as "normal" or "pageLayout".
        /// </summary>
        public string ViewType { get; set; }
        /// <summary>
        /// Gets or sets the top-left visible cell for the worksheet view.
        /// </summary>
        public string TopLeftCell { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether the worksheet tab is selected.
        /// </summary>
        public bool TabSelected { get; set; }
        /// <summary>
        /// Gets or sets the active cell stored in the primary selection.
        /// </summary>
        public string SelectionActiveCell { get; set; }
        /// <summary>
        /// Gets or sets the sqref stored in the primary selection.
        /// </summary>
        public string SelectionSqref { get; set; }
    }
}

