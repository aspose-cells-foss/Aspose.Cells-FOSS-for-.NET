using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents auto filter sort condition model.
    /// </summary>
    public sealed class AutoFilterSortConditionModel
    {
        /// <summary>
        /// Gets or sets the ref.
        /// </summary>
        public string Ref { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether descending.
        /// </summary>
        public bool Descending { get; set; }
        /// <summary>
        /// Gets or sets the sort by.
        /// </summary>
        public string SortBy { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the custom list.
        /// </summary>
        public string CustomList { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the differential style id.
        /// </summary>
        public int? DifferentialStyleId { get; set; }
        /// <summary>
        /// Gets or sets the icon set.
        /// </summary>
        public string IconSet { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the icon id.
        /// </summary>
        public int? IconId { get; set; }
    }
}
