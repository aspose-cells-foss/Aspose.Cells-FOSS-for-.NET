using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class AutoFilter
{
    private readonly AutoFilterModel _model;
    private readonly FilterColumnCollection _filterColumns;
    private readonly AutoFilterSortState _sortState;

    internal AutoFilter(AutoFilterModel model)
    {
        _model = model;
        _filterColumns = new FilterColumnCollection(model.FilterColumns);
        _sortState = new AutoFilterSortState(model.SortState);
    }

    public string Range
    {
        get
        {
            return _model.Range;
        }
        set
        {
            _model.Range = AutoFilterSupport.NormalizeOptionalRange(value, nameof(Range));
        }
    }

    public FilterColumnCollection FilterColumns
    {
        get
        {
            return _filterColumns;
        }
    }

    public AutoFilterSortState SortState
    {
        get
        {
            return _sortState;
        }
    }

    public void Clear()
    {
        _model.Clear();
    }
}

public sealed class FilterColumnCollection : IEnumerable<FilterColumn>
{
    private readonly List<FilterColumnModel> _models;

    internal FilterColumnCollection(List<FilterColumnModel> models)
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

    public FilterColumn this[int index]
    {
        get
        {
            if (index < 0 || index >= _models.Count)
            {
                throw new CellsException("Filter column index was out of range.");
            }

            return new FilterColumn(_models[index]);
        }
    }

    public int Add(int columnIndex)
    {
        if (columnIndex < 0)
        {
            throw new CellsException("Filter column index must be zero or greater.");
        }

        for (var index = 0; index < _models.Count; index++)
        {
            if (_models[index].ColumnIndex == columnIndex)
            {
                throw new CellsException("A filter column for the specified column index already exists.");
            }
        }

        var model = new FilterColumnModel();
        model.ColumnIndex = columnIndex;

        var insertIndex = 0;
        while (insertIndex < _models.Count && _models[insertIndex].ColumnIndex < columnIndex)
        {
            insertIndex++;
        }

        _models.Insert(insertIndex, model);
        return insertIndex;
    }

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= _models.Count)
        {
            throw new CellsException("Filter column index was out of range.");
        }

        _models.RemoveAt(index);
    }

    public void Clear()
    {
        _models.Clear();
    }

    public IEnumerator<FilterColumn> GetEnumerator()
    {
        var columns = new List<FilterColumn>(_models.Count);
        for (var index = 0; index < _models.Count; index++)
        {
            columns.Add(new FilterColumn(_models[index]));
        }

        return columns.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

public sealed class FilterColumn
{
    private readonly FilterColumnModel _model;
    private readonly FilterValueCollection _filters;
    private readonly AutoFilterCustomFilterCollection _customFilters;
    private readonly AutoFilterColorFilter _colorFilter;
    private readonly AutoFilterDynamicFilter _dynamicFilter;
    private readonly AutoFilterTop10 _top10;

    internal FilterColumn(FilterColumnModel model)
    {
        _model = model;
        _filters = new FilterValueCollection(model.Filters);
        _customFilters = new AutoFilterCustomFilterCollection(model.CustomFilters, model);
        _colorFilter = new AutoFilterColorFilter(model.ColorFilter);
        _dynamicFilter = new AutoFilterDynamicFilter(model.DynamicFilter);
        _top10 = new AutoFilterTop10(model.Top10);
    }

    public int ColumnIndex
    {
        get
        {
            return _model.ColumnIndex;
        }
    }

    public bool HiddenButton
    {
        get
        {
            return _model.HiddenButton;
        }
        set
        {
            _model.HiddenButton = value;
        }
    }

    public FilterValueCollection Filters
    {
        get
        {
            return _filters;
        }
    }

    public AutoFilterCustomFilterCollection CustomFilters
    {
        get
        {
            return _customFilters;
        }
    }

    public AutoFilterColorFilter ColorFilter
    {
        get
        {
            return _colorFilter;
        }
    }

    public AutoFilterDynamicFilter DynamicFilter
    {
        get
        {
            return _dynamicFilter;
        }
    }

    public AutoFilterTop10 Top10
    {
        get
        {
            return _top10;
        }
    }

    public void Clear()
    {
        _model.ClearCriteria();
    }
}

public sealed class FilterValueCollection : IEnumerable<string>
{
    private readonly List<string> _values;

    internal FilterValueCollection(List<string> values)
    {
        _values = values;
    }

    public int Count
    {
        get
        {
            return _values.Count;
        }
    }

    public string this[int index]
    {
        get
        {
            if (index < 0 || index >= _values.Count)
            {
                throw new CellsException("Filter value index was out of range.");
            }

            return _values[index];
        }
    }

    public int Add(string value)
    {
        var normalized = AutoFilterSupport.NormalizeText(value, nameof(value));
        _values.Add(normalized);
        return _values.Count - 1;
    }

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= _values.Count)
        {
            throw new CellsException("Filter value index was out of range.");
        }

        _values.RemoveAt(index);
    }

    public void Clear()
    {
        _values.Clear();
    }

    public IEnumerator<string> GetEnumerator()
    {
        var values = new List<string>(_values.Count);
        for (var index = 0; index < _values.Count; index++)
        {
            values.Add(_values[index]);
        }

        return values.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
