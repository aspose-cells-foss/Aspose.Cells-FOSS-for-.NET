using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class Validation
{
    private readonly IList<ValidationModel> _owner;
    private readonly ValidationModel _model;

    internal Validation(IList<ValidationModel> owner, ValidationModel model)
    {
        _owner = owner;
        _model = model;
    }

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

    public void AddArea(CellArea area)
    {
        ValidationCollection.AddAreaToValidation(_owner, _model, area);
    }

    public void RemoveArea(CellArea area)
    {
        ValidationCollection.RemoveAreaFromValidation(_owner, _model, area);
    }

    private static string? NormalizeFormula(string? value)
    {
        if (value is null)
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

    private static string? NormalizeText(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        return value;
    }
}

