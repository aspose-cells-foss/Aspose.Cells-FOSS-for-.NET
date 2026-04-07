using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class FormatCondition
{
    private readonly List<ConditionalFormattingModel> _owner;
    private readonly ConditionalFormattingModel _collection;
    private readonly FormatConditionModel _model;

    internal FormatCondition(List<ConditionalFormattingModel> owner, ConditionalFormattingModel collection, FormatConditionModel model)
    {
        _owner = owner;
        _collection = collection;
        _model = model;
    }

    public FormatConditionType Type
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

    public string Formula
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

    public string TimePeriod
    {
        get
        {
            return _model.TimePeriod ?? string.Empty;
        }
        set
        {
            _model.TimePeriod = NormalizeText(value);
        }
    }

    public bool Duplicate
    {
        get
        {
            return _model.Duplicate;
        }
        set
        {
            _model.Duplicate = value;
        }
    }

    public bool Top
    {
        get
        {
            return _model.Top;
        }
        set
        {
            _model.Top = value;
        }
    }

    public bool Percent
    {
        get
        {
            return _model.Percent;
        }
        set
        {
            _model.Percent = value;
        }
    }

    public int Rank
    {
        get
        {
            return _model.Rank;
        }
        set
        {
            if (value < 0)
            {
                throw new CellsException("Conditional formatting rank must be zero or greater.");
            }

            _model.Rank = value;
        }
    }

    public bool Above
    {
        get
        {
            return _model.Above;
        }
        set
        {
            _model.Above = value;
        }
    }

    public int StandardDeviation
    {
        get
        {
            return _model.StandardDeviation;
        }
        set
        {
            if (value < 0)
            {
                throw new CellsException("Conditional formatting standard deviation must be zero or greater.");
            }

            _model.StandardDeviation = value;
        }
    }

    public int ColorScaleCount
    {
        get
        {
            return _model.ColorScaleCount;
        }
        set
        {
            if (value != 2 && value != 3)
            {
                throw new CellsException("ColorScaleCount must be 2 or 3.");
            }

            _model.ColorScaleCount = value;
        }
    }

    public Color MinColor
    {
        get
        {
            return Color.FromCore(_model.MinColor);
        }
        set
        {
            _model.MinColor = value.ToCore();
        }
    }

    public Color MidColor
    {
        get
        {
            return Color.FromCore(_model.MidColor);
        }
        set
        {
            _model.MidColor = value.ToCore();
        }
    }

    public Color MaxColor
    {
        get
        {
            return Color.FromCore(_model.MaxColor);
        }
        set
        {
            _model.MaxColor = value.ToCore();
        }
    }

    public Color BarColor
    {
        get
        {
            return Color.FromCore(_model.BarColor);
        }
        set
        {
            _model.BarColor = value.ToCore();
        }
    }

    public Color NegativeBarColor
    {
        get
        {
            return Color.FromCore(_model.NegativeBarColor);
        }
        set
        {
            _model.NegativeBarColor = value.ToCore();
        }
    }

    public bool ShowBorder
    {
        get
        {
            return _model.ShowBorder;
        }
        set
        {
            _model.ShowBorder = value;
        }
    }

    public string Direction
    {
        get
        {
            return _model.Direction ?? string.Empty;
        }
        set
        {
            _model.Direction = NormalizeText(value);
        }
    }

    public string BarLength
    {
        get
        {
            return _model.BarLength ?? string.Empty;
        }
        set
        {
            _model.BarLength = NormalizeText(value);
        }
    }

    public string IconSetType
    {
        get
        {
            return _model.IconSetType ?? string.Empty;
        }
        set
        {
            _model.IconSetType = NormalizeText(value);
        }
    }

    public bool ReverseIcons
    {
        get
        {
            return _model.ReverseIcons;
        }
        set
        {
            _model.ReverseIcons = value;
        }
    }

    public bool ShowIconOnly
    {
        get
        {
            return _model.ShowIconOnly;
        }
        set
        {
            _model.ShowIconOnly = value;
        }
    }

    public int Priority
    {
        get
        {
            return _model.Priority;
        }
        set
        {
            if (value <= 0)
            {
                throw new CellsException("Conditional formatting priority must be greater than zero.");
            }

            _model.Priority = value;
        }
    }

    public bool StopIfTrue
    {
        get
        {
            return _model.StopIfTrue;
        }
        set
        {
            _model.StopIfTrue = value;
        }
    }

    public Style Style
    {
        get
        {
            return Style.FromCore(_model.Style).Clone();
        }
        set
        {
            _model.Style = value is null ? StyleValue.Default.Clone() : value.ToCore();
        }
    }

    public void Remove()
    {
        FormatConditionCollection.RemoveCondition(_owner, _collection, _model);
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
        if (value is null)
        {
            return null;
        }

        var trimmed = value.Trim();
        if (trimmed.Length == 0)
        {
            return null;
        }

        return trimmed;
    }
}
