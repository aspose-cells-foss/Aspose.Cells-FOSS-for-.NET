using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public class Worksheet
{
    private readonly Workbook _workbook;
    private readonly WorksheetModel _model;
    private readonly Cells _cells;
    private readonly HyperlinkCollection _hyperlinks;
    private readonly ValidationCollection _validations;
    private readonly ConditionalFormattingCollection _conditionalFormattings;
    private readonly PageSetup _pageSetup;
    private readonly WorksheetProtection _protection;
    private readonly AutoFilter _autoFilter;

    internal Worksheet(Workbook workbook, WorksheetModel model)
    {
        _workbook = workbook;
        _model = model;
        _cells = new Cells(this);
        _hyperlinks = new HyperlinkCollection(model.Hyperlinks);
        _validations = new ValidationCollection(model.Validations);
        _conditionalFormattings = new ConditionalFormattingCollection(model.ConditionalFormattings);
        _pageSetup = new PageSetup(model.PageSetup);
        _protection = new WorksheetProtection(model.Protection);
        _autoFilter = new AutoFilter(model.AutoFilter);
    }

    internal WorksheetModel Model
    {
        get
        {
            return _model;
        }
    }

    internal Workbook Workbook
    {
        get
        {
            return _workbook;
        }
    }

    public string Name
    {
        get
        {
            return _model.Name;
        }
        set
        {
            if (string.IsNullOrWhiteSpace(value)) throw new CellsException("Worksheet name must be non-empty.");
            _workbook.EnsureUniqueSheetName(value, _model);
            _model.Name = value;
        }
    }

    public VisibilityType VisibilityType
    {
        get
        {
            switch (_model.Visibility)
            {
                case SheetVisibility.Hidden:
                    return Aspose.Cells_FOSS.VisibilityType.Hidden;
                case SheetVisibility.VeryHidden:
                    return Aspose.Cells_FOSS.VisibilityType.VeryHidden;
                default:
                    return Aspose.Cells_FOSS.VisibilityType.Visible;
            }
        }
        set
        {
            switch (value)
            {
                case VisibilityType.Hidden:
                    _model.Visibility = SheetVisibility.Hidden;
                    break;
                case VisibilityType.VeryHidden:
                    _model.Visibility = SheetVisibility.VeryHidden;
                    break;
                default:
                    _model.Visibility = SheetVisibility.Visible;
                    break;
            }
        }
    }

    public Color TabColor
    {
        get
        {
            if (_model.TabColor.HasValue)
            {
                return Color.FromCore(_model.TabColor.Value);
            }

            return Color.Empty;
        }
        set
        {
            if (value.Equals(Color.Empty))
            {
                _model.TabColor = null;
                return;
            }

            _model.TabColor = value.ToCore();
        }
    }

    public bool ShowGridlines
    {
        get
        {
            return _model.View.ShowGridLines;
        }
        set
        {
            _model.View.ShowGridLines = value;
        }
    }

    public bool ShowRowColumnHeaders
    {
        get
        {
            return _model.View.ShowRowColumnHeaders;
        }
        set
        {
            _model.View.ShowRowColumnHeaders = value;
        }
    }

    public bool ShowZeros
    {
        get
        {
            return _model.View.ShowZeros;
        }
        set
        {
            _model.View.ShowZeros = value;
        }
    }

    public bool RightToLeft
    {
        get
        {
            return _model.View.RightToLeft;
        }
        set
        {
            _model.View.RightToLeft = value;
        }
    }

    public int Zoom
    {
        get
        {
            return _model.View.ZoomScale;
        }
        set
        {
            if (value < 10 || value > 400)
            {
                throw new CellsException("Zoom must be between 10 and 400.");
            }

            _model.View.ZoomScale = value;
        }
    }

    public Cells Cells
    {
        get
        {
            return _cells;
        }
    }

    public HyperlinkCollection Hyperlinks
    {
        get
        {
            return _hyperlinks;
        }
    }

    public ValidationCollection Validations
    {
        get
        {
            return _validations;
        }
    }

    public ConditionalFormattingCollection ConditionalFormattings
    {
        get
        {
            return _conditionalFormattings;
        }
    }

    public PageSetup PageSetup
    {
        get
        {
            return _pageSetup;
        }
    }

    public WorksheetProtection Protection
    {
        get
        {
            return _protection;
        }
    }

    public AutoFilter AutoFilter
    {
        get
        {
            return _autoFilter;
        }
    }

    public void Protect()
    {
        _model.Protection.IsProtected = true;
    }

    public void Unprotect()
    {
        _protection.Reset();
    }
}
