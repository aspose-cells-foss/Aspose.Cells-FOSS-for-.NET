using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents cell record.
    /// </summary>
    public sealed class CellRecord
    {
        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        public object Value { get; set; }
        /// <summary>
        /// Gets or sets the kind.
        /// </summary>
        public CellValueKind Kind { get; set; } = CellValueKind.Blank;
        /// <summary>
        /// Gets or sets the formula.
        /// </summary>
        public string Formula { get; set; }
        /// <summary>
        /// Gets or sets the style.
        /// </summary>
        public StyleValue Style { get; set; } = StyleValue.Default;
        /// <summary>
        /// Gets or sets a value indicating whether explicitly stored.
        /// </summary>
        public bool IsExplicitlyStored { get; set; }
    }
}
