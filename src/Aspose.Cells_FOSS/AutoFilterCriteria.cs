using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public enum FilterOperatorType
{
    Equal,
    LessThan,
    LessOrEqual,
    NotEqual,
    GreaterOrEqual,
    GreaterThan,
}

public sealed class AutoFilterCustomFilterCollection : IEnumerable<AutoFilterCustomFilter>
{
    private readonly List<AutoFilterCustomFilterModel> _models;
    private readonly FilterColumnModel _columnModel;

    internal AutoFilterCustomFilterCollection(List<AutoFilterCustomFilterModel> models, FilterColumnModel columnModel)
    {
        _models = models;
        _columnModel = columnModel;
    }

    public bool MatchAll
    {
        get
        {
            return _columnModel.CustomFiltersAnd;
        }
        set
        {
            _columnModel.CustomFiltersAnd = value;
        }
    }

    public int Count
    {
        get
        {
            return _models.Count;
        }
    }

    public AutoFilterCustomFilter this[int index]
    {
        get
        {
            if (index < 0 || index >= _models.Count)
            {
                throw new CellsException("Custom filter index was out of range.");
            }

            return new AutoFilterCustomFilter(_models[index]);
        }
    }

    public int Add(FilterOperatorType operatorType, string value)
    {
        if (_models.Count >= 2)
        {
            throw new CellsException("Custom filters support at most two filter conditions.");
        }

        var model = new AutoFilterCustomFilterModel
        {
            Operator = AutoFilterSupport.ToOperatorName(operatorType) ?? string.Empty,
            Value = AutoFilterSupport.NormalizeText(value, nameof(value)),
        };
        _models.Add(model);
        return _models.Count - 1;
    }

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= _models.Count)
        {
            throw new CellsException("Custom filter index was out of range.");
        }

        _models.RemoveAt(index);
    }

    public void Clear()
    {
        _models.Clear();
        _columnModel.CustomFiltersAnd = false;
    }

    public IEnumerator<AutoFilterCustomFilter> GetEnumerator()
    {
        var filters = new List<AutoFilterCustomFilter>(_models.Count);
        for (var index = 0; index < _models.Count; index++)
        {
            filters.Add(new AutoFilterCustomFilter(_models[index]));
        }

        return filters.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

public sealed class AutoFilterCustomFilter
{
    private readonly AutoFilterCustomFilterModel _model;

    internal AutoFilterCustomFilter(AutoFilterCustomFilterModel model)
    {
        _model = model;
    }

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

public sealed class AutoFilterColorFilter
{
    private readonly AutoFilterColorFilterModel _model;

    internal AutoFilterColorFilter(AutoFilterColorFilterModel model)
    {
        _model = model;
    }

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

    public void Clear()
    {
        _model.Clear();
    }
}

public sealed class AutoFilterDynamicFilter
{
    private readonly AutoFilterDynamicFilterModel _model;

    internal AutoFilterDynamicFilter(AutoFilterDynamicFilterModel model)
    {
        _model = model;
    }

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

    public void Clear()
    {
        _model.Clear();
    }
}

public sealed class AutoFilterTop10
{
    private readonly AutoFilterTop10Model _model;

    internal AutoFilterTop10(AutoFilterTop10Model model)
    {
        _model = model;
    }

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

    public void Clear()
    {
        _model.Clear();
    }
}

public sealed class AutoFilterSortState
{
    private readonly AutoFilterSortStateModel _model;
    private readonly AutoFilterSortConditionCollection _conditions;

    internal AutoFilterSortState(AutoFilterSortStateModel model)
    {
        _model = model;
        _conditions = new AutoFilterSortConditionCollection(model.Conditions);
    }

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

    public AutoFilterSortConditionCollection SortConditions
    {
        get
        {
            return _conditions;
        }
    }

    public void Clear()
    {
        _model.Clear();
    }
}

public sealed class AutoFilterSortConditionCollection : IEnumerable<AutoFilterSortCondition>
{
    private readonly List<AutoFilterSortConditionModel> _models;

    internal AutoFilterSortConditionCollection(List<AutoFilterSortConditionModel> models)
    {
        _models = models;
    }

    public int Count
    {
        get
        {
            return _models.Count;
        }
    }

    public AutoFilterSortCondition this[int index]
    {
        get
        {
            if (index < 0 || index >= _models.Count)
            {
                throw new CellsException("Sort condition index was out of range.");
            }

            return new AutoFilterSortCondition(_models[index]);
        }
    }

    public int Add(string reference)
    {
        var model = new AutoFilterSortConditionModel
        {
            Ref = AutoFilterSupport.NormalizeRequiredRange(reference, nameof(reference)),
        };
        _models.Add(model);
        return _models.Count - 1;
    }

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= _models.Count)
        {
            throw new CellsException("Sort condition index was out of range.");
        }

        _models.RemoveAt(index);
    }

    public void Clear()
    {
        _models.Clear();
    }

    public IEnumerator<AutoFilterSortCondition> GetEnumerator()
    {
        var conditions = new List<AutoFilterSortCondition>(_models.Count);
        for (var index = 0; index < _models.Count; index++)
        {
            conditions.Add(new AutoFilterSortCondition(_models[index]));
        }

        return conditions.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

public sealed class AutoFilterSortCondition
{
    private readonly AutoFilterSortConditionModel _model;

    internal AutoFilterSortCondition(AutoFilterSortConditionModel model)
    {
        _model = model;
    }

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
