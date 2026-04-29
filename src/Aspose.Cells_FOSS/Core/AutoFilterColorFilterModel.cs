using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents auto filter color filter model.
    /// </summary>
    public sealed class AutoFilterColorFilterModel
    {
        /// <summary>
        /// Gets or sets a value indicating whether enabled.
        /// </summary>
        public bool Enabled { get; set; }
        /// <summary>
        /// Gets or sets the differential style id.
        /// </summary>
        public int? DifferentialStyleId { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether cell color.
        /// </summary>
        public bool CellColor { get; set; }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            Enabled = false;
            DifferentialStyleId = null;
            CellColor = false;
        }
    }
}
