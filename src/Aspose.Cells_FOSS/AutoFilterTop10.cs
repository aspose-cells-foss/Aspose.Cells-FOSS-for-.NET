using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents auto filter top10.
    /// </summary>
    public sealed class AutoFilterTop10
    {
        private readonly AutoFilterTop10Model _model;

        internal AutoFilterTop10(AutoFilterTop10Model model)
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
        /// Gets or sets a value indicating whether top.
        /// </summary>
        public bool Top
        {
            get
            {
                return _model.Top;
            }
            set
            {
                _model.Top = value;
                _model.Enabled = true;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether percent.
        /// </summary>
        public bool Percent
        {
            get
            {
                return _model.Percent;
            }
            set
            {
                _model.Percent = value;
                _model.Enabled = true;
            }
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        public double? Value
        {
            get
            {
                return _model.Value;
            }
            set
            {
                _model.Value = value;
                if (value.HasValue)
                {
                    _model.Enabled = true;
                }
            }
        }

        /// <summary>
        /// Gets or sets the filter value.
        /// </summary>
        public double? FilterValue
        {
            get
            {
                return _model.FilterValue;
            }
            set
            {
                _model.FilterValue = value;
                if (value.HasValue)
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
