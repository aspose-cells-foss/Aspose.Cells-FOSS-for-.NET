using System.Linq;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents auto filter.
    /// </summary>
    public sealed class AutoFilter
    {
        private readonly AutoFilterModel _model;
        private readonly FilterColumnCollection _filterColumns;
        private readonly AutoFilterSortState _sortState;

        internal AutoFilter(AutoFilterModel model)
        {
            _model = model;
            _filterColumns = new FilterColumnCollection(model.FilterColumns);
            _sortState = new AutoFilterSortState(model.SortState);
        }

        /// <summary>
        /// Gets or sets the range.
        /// </summary>
        public string Range
        {
            get
            {
                return _model.Range;
            }
            set
            {
                _model.Range = AutoFilterSupport.NormalizeOptionalRange(value, nameof(Range));
            }
        }

        /// <summary>
        /// Gets the filter columns.
        /// </summary>
        public FilterColumnCollection FilterColumns
        {
            get
            {
                return _filterColumns;
            }
        }

        /// <summary>
        /// Gets the sort state.
        /// </summary>
        public AutoFilterSortState SortState
        {
            get
            {
                return _sortState;
            }
        }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            _model.Clear();
        }
    }
}
