using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents filter column.
    /// </summary>
    public sealed class FilterColumn
    {
        private readonly FilterColumnModel _model;
        private readonly FilterValueCollection _filters;
        private readonly AutoFilterCustomFilterCollection _customFilters;
        private readonly AutoFilterColorFilter _colorFilter;
        private readonly AutoFilterDynamicFilter _dynamicFilter;
        private readonly AutoFilterTop10 _top10;

        internal FilterColumn(FilterColumnModel model)
        {
            _model = model;
            _filters = new FilterValueCollection(model.Filters);
            _customFilters = new AutoFilterCustomFilterCollection(model.CustomFilters, model);
            _colorFilter = new AutoFilterColorFilter(model.ColorFilter);
            _dynamicFilter = new AutoFilterDynamicFilter(model.DynamicFilter);
            _top10 = new AutoFilterTop10(model.Top10);
        }

        /// <summary>
        /// Gets the column index.
        /// </summary>
        public int ColumnIndex
        {
            get
            {
                return _model.ColumnIndex;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether hidden button.
        /// </summary>
        public bool HiddenButton
        {
            get
            {
                return _model.HiddenButton;
            }
            set
            {
                _model.HiddenButton = value;
            }
        }

        /// <summary>
        /// Gets the filters.
        /// </summary>
        public FilterValueCollection Filters
        {
            get
            {
                return _filters;
            }
        }

        /// <summary>
        /// Gets the custom filters.
        /// </summary>
        public AutoFilterCustomFilterCollection CustomFilters
        {
            get
            {
                return _customFilters;
            }
        }

        /// <summary>
        /// Gets the color filter.
        /// </summary>
        public AutoFilterColorFilter ColorFilter
        {
            get
            {
                return _colorFilter;
            }
        }

        /// <summary>
        /// Gets the dynamic filter.
        /// </summary>
        public AutoFilterDynamicFilter DynamicFilter
        {
            get
            {
                return _dynamicFilter;
            }
        }

        /// <summary>
        /// Gets the top10.
        /// </summary>
        public AutoFilterTop10 Top10
        {
            get
            {
                return _top10;
            }
        }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            _model.ClearCriteria();
        }
    }
}
