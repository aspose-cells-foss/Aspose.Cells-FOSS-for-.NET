using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents auto filter sort state model.
    /// </summary>
    public sealed class AutoFilterSortStateModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AutoFilterSortStateModel"/> class.
        /// </summary>
        public AutoFilterSortStateModel()
        {
            Ref = string.Empty;
            SortMethod = string.Empty;
            Conditions = new List<AutoFilterSortConditionModel>();
        }

        /// <summary>
        /// Gets or sets a value indicating whether column sort.
        /// </summary>
        public bool ColumnSort { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether case sensitive.
        /// </summary>
        public bool CaseSensitive { get; set; }
        /// <summary>
        /// Gets or sets the sort method.
        /// </summary>
        public string SortMethod { get; set; }
        /// <summary>
        /// Gets or sets the ref.
        /// </summary>
        public string Ref { get; set; }
        /// <summary>
        /// Gets the conditions.
        /// </summary>
        public List<AutoFilterSortConditionModel> Conditions { get; }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            ColumnSort = false;
            CaseSensitive = false;
            SortMethod = string.Empty;
            Ref = string.Empty;
            Conditions.Clear();
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return !string.IsNullOrEmpty(Ref)
                || ColumnSort
                || CaseSensitive
                || !string.IsNullOrEmpty(SortMethod)
                || Conditions.Count > 0;
        }
    }
}
