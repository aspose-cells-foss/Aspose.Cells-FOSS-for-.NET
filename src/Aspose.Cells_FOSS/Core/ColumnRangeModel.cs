using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents column range model.
    /// </summary>
    public sealed class ColumnRangeModel
    {
        /// <summary>
        /// Gets or sets the min column index.
        /// </summary>
        public int MinColumnIndex { get; set; }
        /// <summary>
        /// Gets or sets the max column index.
        /// </summary>
        public int MaxColumnIndex { get; set; }
        /// <summary>
        /// Gets or sets the width.
        /// </summary>
        public double? Width { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether hidden.
        /// </summary>
        public bool Hidden { get; set; }
        /// <summary>
        /// Gets or sets the style index.
        /// </summary>
        public int? StyleIndex { get; set; }
    }
}
