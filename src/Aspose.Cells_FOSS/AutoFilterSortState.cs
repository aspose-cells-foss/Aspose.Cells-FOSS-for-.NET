using System.Linq;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents auto filter sort state.
    /// </summary>
    public sealed class AutoFilterSortState
    {
        private readonly AutoFilterSortStateModel _model;
        private readonly AutoFilterSortConditionCollection _conditions;

        internal AutoFilterSortState(AutoFilterSortStateModel model)
        {
            _model = model;
            _conditions = new AutoFilterSortConditionCollection(model.Conditions);
        }

        /// <summary>
        /// Gets or sets a value indicating whether column sort.
        /// </summary>
        public bool ColumnSort
        {
            get
            {
                return _model.ColumnSort;
            }
            set
            {
                _model.ColumnSort = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether case sensitive.
        /// </summary>
        public bool CaseSensitive
        {
            get
            {
                return _model.CaseSensitive;
            }
            set
            {
                _model.CaseSensitive = value;
            }
        }

        /// <summary>
        /// Gets or sets the sort method.
        /// </summary>
        public string SortMethod
        {
            get
            {
                return _model.SortMethod;
            }
            set
            {
                _model.SortMethod = AutoFilterSupport.NormalizeOptionalText(value);
            }
        }

        /// <summary>
        /// Gets or sets the ref.
        /// </summary>
        public string Ref
        {
            get
            {
                return _model.Ref;
            }
            set
            {
                _model.Ref = AutoFilterSupport.NormalizeOptionalRange(value, nameof(Ref));
            }
        }

        /// <summary>
        /// Gets the sort conditions.
        /// </summary>
        public AutoFilterSortConditionCollection SortConditions
        {
            get
            {
                return _conditions;
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
