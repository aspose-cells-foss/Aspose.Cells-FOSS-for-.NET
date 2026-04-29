using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents auto filter custom filter.
    /// </summary>
    public sealed class AutoFilterCustomFilter
    {
        private readonly AutoFilterCustomFilterModel _model;

        internal AutoFilterCustomFilter(AutoFilterCustomFilterModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets the operator.
        /// </summary>
        public FilterOperatorType Operator
        {
            get
            {
                return AutoFilterSupport.ParseOperatorOrDefault(_model.Operator);
            }
            set
            {
                _model.Operator = AutoFilterSupport.ToOperatorName(value) ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        public string Value
        {
            get
            {
                return _model.Value;
            }
            set
            {
                _model.Value = AutoFilterSupport.NormalizeText(value, nameof(Value));
            }
        }
    }
}
