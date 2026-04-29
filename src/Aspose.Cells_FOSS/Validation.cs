using System.IO;
using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents validation.
    /// </summary>
    public sealed class Validation
    {
        private readonly IList<ValidationModel> _owner;
        private readonly ValidationModel _model;

        internal Validation(IList<ValidationModel> owner, ValidationModel model)
        {
            _owner = owner;
            _model = model;
        }

        /// <summary>
        /// Gets the areas.
        /// </summary>
        public IReadOnlyList<CellArea> Areas
        {
            get
            {
                var areas = new List<CellArea>(_model.Areas.Count);
                for (var index = 0; index < _model.Areas.Count; index++)
                {
                    areas.Add(_model.Areas[index]);
                }

                return areas;
            }
        }

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        public ValidationType Type
        {
            get
            {
                return _model.Type;
            }
            set
            {
                _model.Type = value;
            }
        }

        /// <summary>
        /// Gets or sets the alert style.
        /// </summary>
        public ValidationAlertType AlertStyle
        {
            get
            {
                return _model.AlertStyle;
            }
            set
            {
                _model.AlertStyle = value;
            }
        }

        /// <summary>
        /// Gets or sets the operator.
        /// </summary>
        public OperatorType Operator
        {
            get
            {
                return _model.Operator;
            }
            set
            {
                _model.Operator = value;
            }
        }

        /// <summary>
        /// Gets or sets the formula1.
        /// </summary>
        public string Formula1
        {
            get
            {
                return _model.Formula1 ?? string.Empty;
            }
            set
            {
                _model.Formula1 = NormalizeFormula(value);
            }
        }

        /// <summary>
        /// Gets or sets the formula2.
        /// </summary>
        public string Formula2
        {
            get
            {
                return _model.Formula2 ?? string.Empty;
            }
            set
            {
                _model.Formula2 = NormalizeFormula(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether ignore blank.
        /// </summary>
        public bool IgnoreBlank
        {
            get
            {
                return _model.IgnoreBlank;
            }
            set
            {
                _model.IgnoreBlank = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether in cell drop down.
        /// </summary>
        public bool InCellDropDown
        {
            get
            {
                return _model.InCellDropDown;
            }
            set
            {
                _model.InCellDropDown = value;
            }
        }

        /// <summary>
        /// Gets or sets the input title.
        /// </summary>
        public string InputTitle
        {
            get
            {
                return _model.InputTitle ?? string.Empty;
            }
            set
            {
                _model.InputTitle = NormalizeText(value);
            }
        }

        /// <summary>
        /// Gets or sets the input message.
        /// </summary>
        public string InputMessage
        {
            get
            {
                return _model.InputMessage ?? string.Empty;
            }
            set
            {
                _model.InputMessage = NormalizeText(value);
            }
        }

        /// <summary>
        /// Gets or sets the error title.
        /// </summary>
        public string ErrorTitle
        {
            get
            {
                return _model.ErrorTitle ?? string.Empty;
            }
            set
            {
                _model.ErrorTitle = NormalizeText(value);
            }
        }

        /// <summary>
        /// Gets or sets the error message.
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return _model.ErrorMessage ?? string.Empty;
            }
            set
            {
                _model.ErrorMessage = NormalizeText(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether show input.
        /// </summary>
        public bool ShowInput
        {
            get
            {
                return _model.ShowInput;
            }
            set
            {
                _model.ShowInput = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether show error.
        /// </summary>
        public bool ShowError
        {
            get
            {
                return _model.ShowError;
            }
            set
            {
                _model.ShowError = value;
            }
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <param name="area">The area.</param>
        public void AddArea(CellArea area)
        {
            ValidationCollection.AddAreaToValidation(_owner, _model, area);
        }

        /// <summary>
        /// Removes the specified item.
        /// </summary>
        /// <param name="area">The area.</param>
        public void RemoveArea(CellArea area)
        {
            ValidationCollection.RemoveAreaFromValidation(_owner, _model, area);
        }

        private static string NormalizeFormula(string value)
        {
            if (value == null)
            {
                return null;
            }

            var trimmed = value.Trim();
            if (trimmed.Length == 0)
            {
                return null;
            }

            if (trimmed[0] == '=')
            {
                return trimmed.Substring(1);
            }

            return trimmed;
        }

        private static string NormalizeText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }

            return value;
        }
    }
}
