using System.Linq;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents auto filter color filter.
    /// </summary>
    public sealed class AutoFilterColorFilter
    {
        private readonly AutoFilterColorFilterModel _model;

        internal AutoFilterColorFilter(AutoFilterColorFilterModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets a value indicating whether enabled.
        /// </summary>
        public bool Enabled
        {
            get
            {
                return _model.Enabled;
            }
            set
            {
                if (!value)
                {
                    _model.Clear();
                    return;
                }

                _model.Enabled = true;
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
                if (value.HasValue)
                {
                    _model.Enabled = true;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether cell color.
        /// </summary>
        public bool CellColor
        {
            get
            {
                return _model.CellColor;
            }
            set
            {
                _model.CellColor = value;
                if (value)
                {
                    _model.Enabled = true;
                }
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
