using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents auto filter dynamic filter model.
    /// </summary>
    public sealed class AutoFilterDynamicFilterModel
    {
        /// <summary>
        /// Gets or sets a value indicating whether enabled.
        /// </summary>
        public bool Enabled { get; set; }
        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        public double? Value { get; set; }
        /// <summary>
        /// Gets or sets the max value.
        /// </summary>
        public double? MaxValue { get; set; }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            Enabled = false;
            Type = string.Empty;
            Value = null;
            MaxValue = null;
        }
    }
}
