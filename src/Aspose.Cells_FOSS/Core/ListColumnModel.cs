using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents the internal model for a single column in a table.
    /// </summary>
    public sealed class ListColumnModel
    {
        /// <summary>
        /// Initializes a new instance with the given one-based id and name.
        /// </summary>
        public ListColumnModel(int id, string name)
        {
            Id = id;
            Name = name;
            TotalsRowFunction = "none";
            TotalsRowLabel = string.Empty;
            TotalsRowFormula = string.Empty;
        }

        /// <summary>
        /// Gets the one-based column id within the table.
        /// </summary>
        public int Id { get; }

        /// <summary>
        /// Gets or sets the column header name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the SpreadsheetML totalsRowFunction attribute value.
        /// </summary>
        public string TotalsRowFunction { get; set; }

        /// <summary>
        /// Gets or sets the label shown in the totals row cell for this column.
        /// </summary>
        public string TotalsRowLabel { get; set; }

        /// <summary>
        /// Gets or sets the custom formula text used when TotalsRowFunction is "custom".
        /// </summary>
        public string TotalsRowFormula { get; set; }
    }
}
