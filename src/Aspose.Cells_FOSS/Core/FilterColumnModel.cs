using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents filter column model.
    /// </summary>
    public sealed class FilterColumnModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FilterColumnModel"/> class.
        /// </summary>
        public FilterColumnModel()
        {
            Filters = new List<string>();
            CustomFilters = new List<AutoFilterCustomFilterModel>();
            ColorFilter = new AutoFilterColorFilterModel();
            DynamicFilter = new AutoFilterDynamicFilterModel();
            Top10 = new AutoFilterTop10Model();
        }

        /// <summary>
        /// Gets or sets the column index.
        /// </summary>
        public int ColumnIndex { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether hidden button.
        /// </summary>
        public bool HiddenButton { get; set; }
        /// <summary>
        /// Gets the filters.
        /// </summary>
        public List<string> Filters { get; }
        /// <summary>
        /// Gets or sets the custom filters.
        /// </summary>
        public List<AutoFilterCustomFilterModel> CustomFilters { get; }
        /// <summary>
        /// Gets or sets a value indicating whether custom filters and.
        /// </summary>
        public bool CustomFiltersAnd { get; set; }
        /// <summary>
        /// Gets the color filter.
        /// </summary>
        public AutoFilterColorFilterModel ColorFilter { get; }
        /// <summary>
        /// Gets the dynamic filter.
        /// </summary>
        public AutoFilterDynamicFilterModel DynamicFilter { get; }
        /// <summary>
        /// Gets the top10.
        /// </summary>
        public AutoFilterTop10Model Top10 { get; }

        /// <summary>
        /// Performs clear criteria.
        /// </summary>
        public void ClearCriteria()
        {
            HiddenButton = false;
            Filters.Clear();
            CustomFilters.Clear();
            CustomFiltersAnd = false;
            ColorFilter.Clear();
            DynamicFilter.Clear();
            Top10.Clear();
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return HiddenButton
                || Filters.Count > 0
                || CustomFilters.Count > 0
                || ColorFilter.Enabled
                || DynamicFilter.Enabled
                || Top10.Enabled;
        }
    }
}
