using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class AutoFilterModel
{
    public AutoFilterModel()
    {
        Range = string.Empty;
        FilterColumns = new List<FilterColumnModel>();
        SortState = new AutoFilterSortStateModel();
    }

    public string Range { get; set; }
    public List<FilterColumnModel> FilterColumns { get; }
    public AutoFilterSortStateModel SortState { get; }

    public void Clear()
    {
        Range = string.Empty;
        FilterColumns.Clear();
        SortState.Clear();
    }

    public bool HasStoredState()
    {
        return !string.IsNullOrEmpty(Range);
    }
}

public sealed class FilterColumnModel
{
    public FilterColumnModel()
    {
        Filters = new List<string>();
        CustomFilters = new List<AutoFilterCustomFilterModel>();
        ColorFilter = new AutoFilterColorFilterModel();
        DynamicFilter = new AutoFilterDynamicFilterModel();
        Top10 = new AutoFilterTop10Model();
    }

    public int ColumnIndex { get; set; }
    public bool HiddenButton { get; set; }
    public List<string> Filters { get; }
    public List<AutoFilterCustomFilterModel> CustomFilters { get; }
    public bool CustomFiltersAnd { get; set; }
    public AutoFilterColorFilterModel ColorFilter { get; }
    public AutoFilterDynamicFilterModel DynamicFilter { get; }
    public AutoFilterTop10Model Top10 { get; }

    public void ClearCriteria()
    {
        HiddenButton = false;
        Filters.Clear();
        CustomFilters.Clear();
        CustomFiltersAnd = false;
        ColorFilter.Clear();
        DynamicFilter.Clear();
        Top10.Clear();
    }

    public bool HasStoredState()
    {
        return HiddenButton
            || Filters.Count > 0
            || CustomFilters.Count > 0
            || ColorFilter.Enabled
            || DynamicFilter.Enabled
            || Top10.Enabled;
    }
}

public sealed class AutoFilterCustomFilterModel
{
    public string Operator { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

public sealed class AutoFilterColorFilterModel
{
    public bool Enabled { get; set; }
    public int? DifferentialStyleId { get; set; }
    public bool CellColor { get; set; }

    public void Clear()
    {
        Enabled = false;
        DifferentialStyleId = null;
        CellColor = false;
    }
}

public sealed class AutoFilterDynamicFilterModel
{
    public bool Enabled { get; set; }
    public string Type { get; set; } = string.Empty;
    public double? Value { get; set; }
    public double? MaxValue { get; set; }

    public void Clear()
    {
        Enabled = false;
        Type = string.Empty;
        Value = null;
        MaxValue = null;
    }
}

public sealed class AutoFilterTop10Model
{
    public bool Enabled { get; set; }
    public bool Top { get; set; } = true;
    public bool Percent { get; set; }
    public double? Value { get; set; }
    public double? FilterValue { get; set; }

    public void Clear()
    {
        Enabled = false;
        Top = true;
        Percent = false;
        Value = null;
        FilterValue = null;
    }
}

public sealed class AutoFilterSortStateModel
{
    public AutoFilterSortStateModel()
    {
        Ref = string.Empty;
        SortMethod = string.Empty;
        Conditions = new List<AutoFilterSortConditionModel>();
    }

    public bool ColumnSort { get; set; }
    public bool CaseSensitive { get; set; }
    public string SortMethod { get; set; }
    public string Ref { get; set; }
    public List<AutoFilterSortConditionModel> Conditions { get; }

    public void Clear()
    {
        ColumnSort = false;
        CaseSensitive = false;
        SortMethod = string.Empty;
        Ref = string.Empty;
        Conditions.Clear();
    }

    public bool HasStoredState()
    {
        return !string.IsNullOrEmpty(Ref)
            || ColumnSort
            || CaseSensitive
            || !string.IsNullOrEmpty(SortMethod)
            || Conditions.Count > 0;
    }
}

public sealed class AutoFilterSortConditionModel
{
    public string Ref { get; set; } = string.Empty;
    public bool Descending { get; set; }
    public string SortBy { get; set; } = string.Empty;
    public string CustomList { get; set; } = string.Empty;
    public int? DifferentialStyleId { get; set; }
    public string IconSet { get; set; } = string.Empty;
    public int? IconId { get; set; }
}
