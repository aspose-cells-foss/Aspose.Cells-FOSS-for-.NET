using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class WorkbookView
{
    private readonly WorkbookModel _workbookModel;
    private readonly WorkbookViewModel _model;

    internal WorkbookView(WorkbookModel workbookModel)
    {
        _workbookModel = workbookModel;
        _model = workbookModel.Properties.View;
    }

    public int XWindow
    {
        get
        {
            return _model.XWindow ?? 0;
        }
        set
        {
            _model.XWindow = value;
        }
    }

    public int YWindow
    {
        get
        {
            return _model.YWindow ?? 0;
        }
        set
        {
            _model.YWindow = value;
        }
    }

    public int WindowWidth
    {
        get
        {
            return _model.WindowWidth ?? 0;
        }
        set
        {
            if (value < 0)
            {
                throw new CellsException("WindowWidth must be non-negative.");
            }

            _model.WindowWidth = value;
        }
    }

    public int WindowHeight
    {
        get
        {
            return _model.WindowHeight ?? 0;
        }
        set
        {
            if (value < 0)
            {
                throw new CellsException("WindowHeight must be non-negative.");
            }

            _model.WindowHeight = value;
        }
    }

    public int ActiveTab
    {
        get
        {
            return _workbookModel.ActiveSheetIndex;
        }
        set
        {
            if (value < 0 || value >= _workbookModel.Worksheets.Count)
            {
                throw new CellsException("ActiveTab must refer to an existing worksheet.");
            }

            _workbookModel.ActiveSheetIndex = value;
        }
    }

    public int FirstSheet
    {
        get
        {
            return _model.FirstSheet ?? 0;
        }
        set
        {
            if (value < 0)
            {
                throw new CellsException("FirstSheet must be non-negative.");
            }

            _model.FirstSheet = value;
        }
    }

    public bool ShowHorizontalScroll
    {
        get
        {
            return !_model.ShowHorizontalScroll.HasValue || _model.ShowHorizontalScroll.Value;
        }
        set
        {
            _model.ShowHorizontalScroll = value;
        }
    }

    public bool ShowVerticalScroll
    {
        get
        {
            return !_model.ShowVerticalScroll.HasValue || _model.ShowVerticalScroll.Value;
        }
        set
        {
            _model.ShowVerticalScroll = value;
        }
    }

    public bool ShowSheetTabs
    {
        get
        {
            return !_model.ShowSheetTabs.HasValue || _model.ShowSheetTabs.Value;
        }
        set
        {
            _model.ShowSheetTabs = value;
        }
    }

    public int TabRatio
    {
        get
        {
            return _model.TabRatio ?? 600;
        }
        set
        {
            if (value < 0 || value > 1000)
            {
                throw new CellsException("TabRatio must be between 0 and 1000.");
            }

            _model.TabRatio = value;
        }
    }

    public string Visibility
    {
        get
        {
            return string.IsNullOrEmpty(_model.Visibility) ? "visible" : _model.Visibility;
        }
        set
        {
            _model.Visibility = WorkbookPropertySupport.NormalizeVisibility(value);
        }
    }

    public bool Minimized
    {
        get
        {
            return _model.Minimized;
        }
        set
        {
            _model.Minimized = value;
        }
    }

    public bool AutoFilterDateGrouping
    {
        get
        {
            return _model.AutoFilterDateGrouping;
        }
        set
        {
            _model.AutoFilterDateGrouping = value;
        }
    }
}
