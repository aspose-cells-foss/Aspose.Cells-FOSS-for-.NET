using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents the internal model for an Excel table (ListObject).
    /// StartRow/StartColumn/EndRow/EndColumn are zero-based inclusive bounds
    /// covering the full table range including header and totals rows.
    /// </summary>
    public sealed class ListObjectModel
    {
        /// <summary>
        /// Initializes a new instance with an empty column list.
        /// </summary>
        public ListObjectModel()
        {
            Columns = new List<ListColumnModel>();
            DisplayName = string.Empty;
            Name = string.Empty;
            Comment = string.Empty;
            TableStyleName = string.Empty;
            ShowHeaderRow = true;
            ShowRowStripes = true;
        }

        /// <summary>
        /// Gets or sets the user-visible table name (no spaces allowed).
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets the SpreadsheetML name attribute (mirrors DisplayName).
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the optional comment for the table.
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// Gets or sets the zero-based index of the first row (header or first data row).
        /// </summary>
        public int StartRow { get; set; }

        /// <summary>
        /// Gets or sets the zero-based index of the first column.
        /// </summary>
        public int StartColumn { get; set; }

        /// <summary>
        /// Gets or sets the zero-based index of the last row (last data or totals row).
        /// </summary>
        public int EndRow { get; set; }

        /// <summary>
        /// Gets or sets the zero-based index of the last column.
        /// </summary>
        public int EndColumn { get; set; }

        /// <summary>
        /// Gets or sets whether the first row of the range is a header row.
        /// </summary>
        public bool ShowHeaderRow { get; set; }

        /// <summary>
        /// Gets or sets whether the last row of the range is a totals row.
        /// </summary>
        public bool ShowTotals { get; set; }

        /// <summary>
        /// Gets or sets whether the autoFilter element is emitted inside the table part.
        /// </summary>
        public bool HasAutoFilter { get; set; }

        /// <summary>
        /// Gets or sets the built-in or custom table style name.
        /// </summary>
        public string TableStyleName { get; set; }

        /// <summary>
        /// Gets or sets whether the first column receives special styling.
        /// </summary>
        public bool ShowFirstColumn { get; set; }

        /// <summary>
        /// Gets or sets whether the last column receives special styling.
        /// </summary>
        public bool ShowLastColumn { get; set; }

        /// <summary>
        /// Gets or sets whether row banding stripes are shown.
        /// </summary>
        public bool ShowRowStripes { get; set; }

        /// <summary>
        /// Gets or sets whether column banding stripes are shown.
        /// </summary>
        public bool ShowColumnStripes { get; set; }

        /// <summary>
        /// Gets the ordered list of column models for this table.
        /// </summary>
        public List<ListColumnModel> Columns { get; }
    }
}
