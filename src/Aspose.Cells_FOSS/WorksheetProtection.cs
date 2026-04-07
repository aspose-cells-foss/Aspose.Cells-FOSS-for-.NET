using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class WorksheetProtection
{
    private readonly WorksheetProtectionModel _model;

    internal WorksheetProtection(WorksheetProtectionModel model)
    {
        _model = model;
    }

    public bool IsProtected
    {
        get
        {
            return _model.IsProtected;
        }
        set
        {
            _model.IsProtected = value;
        }
    }

    public bool Objects
    {
        get
        {
            return _model.Objects;
        }
        set
        {
            _model.Objects = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool Scenarios
    {
        get
        {
            return _model.Scenarios;
        }
        set
        {
            _model.Scenarios = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool FormatCells
    {
        get
        {
            return _model.FormatCells;
        }
        set
        {
            _model.FormatCells = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool FormatColumns
    {
        get
        {
            return _model.FormatColumns;
        }
        set
        {
            _model.FormatColumns = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool FormatRows
    {
        get
        {
            return _model.FormatRows;
        }
        set
        {
            _model.FormatRows = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool InsertColumns
    {
        get
        {
            return _model.InsertColumns;
        }
        set
        {
            _model.InsertColumns = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool InsertRows
    {
        get
        {
            return _model.InsertRows;
        }
        set
        {
            _model.InsertRows = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool InsertHyperlinks
    {
        get
        {
            return _model.InsertHyperlinks;
        }
        set
        {
            _model.InsertHyperlinks = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool DeleteColumns
    {
        get
        {
            return _model.DeleteColumns;
        }
        set
        {
            _model.DeleteColumns = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool DeleteRows
    {
        get
        {
            return _model.DeleteRows;
        }
        set
        {
            _model.DeleteRows = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool SelectLockedCells
    {
        get
        {
            return _model.SelectLockedCells;
        }
        set
        {
            _model.SelectLockedCells = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool Sort
    {
        get
        {
            return _model.Sort;
        }
        set
        {
            _model.Sort = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool AutoFilter
    {
        get
        {
            return _model.AutoFilter;
        }
        set
        {
            _model.AutoFilter = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool PivotTables
    {
        get
        {
            return _model.PivotTables;
        }
        set
        {
            _model.PivotTables = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    public bool SelectUnlockedCells
    {
        get
        {
            return _model.SelectUnlockedCells;
        }
        set
        {
            _model.SelectUnlockedCells = value;
            MarkProtectedWhenEnabled(value);
        }
    }

    internal void Reset()
    {
        _model.Clear();
    }

    private void MarkProtectedWhenEnabled(bool value)
    {
        if (value)
        {
            _model.IsProtected = true;
        }
    }
}
