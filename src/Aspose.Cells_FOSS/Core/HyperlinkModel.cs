using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents hyperlink model.
    /// </summary>
    public sealed class HyperlinkModel
    {
        /// <summary>
        /// Gets or sets the first row.
        /// </summary>
        public int FirstRow { get; set; }
        /// <summary>
        /// Gets or sets the first column.
        /// </summary>
        public int FirstColumn { get; set; }
        /// <summary>
        /// Gets or sets the total rows.
        /// </summary>
        public int TotalRows { get; set; } = 1;
        /// <summary>
        /// Gets or sets the total columns.
        /// </summary>
        public int TotalColumns { get; set; } = 1;
        /// <summary>
        /// Gets or sets the address.
        /// </summary>
        public string Address { get; set; }
        /// <summary>
        /// Gets or sets the sub address.
        /// </summary>
        public string SubAddress { get; set; }
        /// <summary>
        /// Gets or sets the screen tip.
        /// </summary>
        public string ScreenTip { get; set; }
        /// <summary>
        /// Gets or sets the text to display.
        /// </summary>
        public string TextToDisplay { get; set; }
    }
}
