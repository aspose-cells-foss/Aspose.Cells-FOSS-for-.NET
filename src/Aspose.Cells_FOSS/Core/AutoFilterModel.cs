using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents auto filter model.
    /// </summary>
    public sealed class AutoFilterModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AutoFilterModel"/> class.
        /// </summary>
        public AutoFilterModel()
        {
            Range = string.Empty;
            FilterColumns = new List<FilterColumnModel>();
            SortState = new AutoFilterSortStateModel();
        }

        /// <summary>
        /// Gets or sets the range.
        /// </summary>
        public string Range { get; set; }
        /// <summary>
        /// Gets the filter columns.
        /// </summary>
        public List<FilterColumnModel> FilterColumns { get; }
        /// <summary>
        /// Gets the sort state.
        /// </summary>
        public AutoFilterSortStateModel SortState { get; }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            Range = string.Empty;
            FilterColumns.Clear();
            SortState.Clear();
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return !string.IsNullOrEmpty(Range);
        }
    }
}
