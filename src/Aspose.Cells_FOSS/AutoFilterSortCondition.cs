using System.Linq;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents auto filter sort condition.
    /// </summary>
    public sealed class AutoFilterSortCondition
    {
        private readonly AutoFilterSortConditionModel _model;

        internal AutoFilterSortCondition(AutoFilterSortConditionModel model)
        {
            _model = model;
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
                _model.Ref = AutoFilterSupport.NormalizeRequiredRange(value, nameof(Ref));
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether descending.
        /// </summary>
        public bool Descending
        {
            get
            {
                return _model.Descending;
            }
            set
            {
                _model.Descending = value;
            }
        }

        /// <summary>
        /// Gets or sets the sort by.
        /// </summary>
        public string SortBy
        {
            get
            {
                return _model.SortBy;
            }
            set
            {
                _model.SortBy = AutoFilterSupport.NormalizeOptionalText(value);
            }
        }

        /// <summary>
        /// Gets or sets the custom list.
        /// </summary>
        public string CustomList
        {
            get
            {
                return _model.CustomList;
            }
            set
            {
                _model.CustomList = AutoFilterSupport.NormalizeOptionalText(value);
            }
        }

        /// <summary>
        /// Gets or sets the differential style id.
        /// </summary>
        public int? DifferentialStyleId
        {
            get
            {
                return _model.DifferentialStyleId;
            }
            set
            {
                if (value.HasValue && value.Value < 0)
                {
                    throw new CellsException("Differential style id must be zero or greater.");
                }

                _model.DifferentialStyleId = value;
            }
        }

        /// <summary>
        /// Gets or sets the icon set.
        /// </summary>
        public string IconSet
        {
            get
            {
                return _model.IconSet;
            }
            set
            {
                _model.IconSet = AutoFilterSupport.NormalizeOptionalText(value);
            }
        }

        /// <summary>
        /// Gets or sets the icon id.
        /// </summary>
        public int? IconId
        {
            get
            {
                return _model.IconId;
            }
            set
            {
                if (value.HasValue && value.Value < 0)
                {
                    throw new CellsException("Icon id must be zero or greater.");
                }

                _model.IconId = value;
            }
        }
    }
}
