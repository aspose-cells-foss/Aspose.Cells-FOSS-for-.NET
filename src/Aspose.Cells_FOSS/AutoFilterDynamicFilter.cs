using System.Linq;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents auto filter dynamic filter.
    /// </summary>
    public sealed class AutoFilterDynamicFilter
    {
        private readonly AutoFilterDynamicFilterModel _model;

        internal AutoFilterDynamicFilter(AutoFilterDynamicFilterModel model)
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
        /// Gets or sets the type.
        /// </summary>
        public string Type
        {
            get
            {
                return _model.Type;
            }
            set
            {
                _model.Type = AutoFilterSupport.NormalizeOptionalText(value);
                if (_model.Type.Length > 0)
                {
                    _model.Enabled = true;
                }
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
        /// Gets or sets the max value.
        /// </summary>
        public double? MaxValue
        {
            get
            {
                return _model.MaxValue;
            }
            set
            {
                _model.MaxValue = value;
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
