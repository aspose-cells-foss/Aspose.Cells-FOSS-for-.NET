using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents auto filter top10 model.
    /// </summary>
    public sealed class AutoFilterTop10Model
    {
        /// <summary>
        /// Gets or sets a value indicating whether enabled.
        /// </summary>
        public bool Enabled { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether top.
        /// </summary>
        public bool Top { get; set; } = true;
        /// <summary>
        /// Gets or sets a value indicating whether percent.
        /// </summary>
        public bool Percent { get; set; }
        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        public double? Value { get; set; }
        /// <summary>
        /// Gets or sets the filter value.
        /// </summary>
        public double? FilterValue { get; set; }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            Enabled = false;
            Top = true;
            Percent = false;
            Value = null;
            FilterValue = null;
        }
    }
}
