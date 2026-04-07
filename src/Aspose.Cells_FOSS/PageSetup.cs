using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class PageSetup
{
    private const double CentimetersPerInch = 2.54d;
    private readonly PageSetupModel _model;

    internal PageSetup(PageSetupModel model)
    {
        _model = model;
    }

    public PaperSizeType PaperSize
    {
        get
        {
            return (PaperSizeType)_model.PaperSize;
        }
        set
        {
            _model.PaperSize = (int)value;
        }
    }

    public PageOrientationType Orientation
    {
        get
        {
            switch (_model.Orientation)
            {
                case PageOrientation.Portrait:
                    return PageOrientationType.Portrait;
                case PageOrientation.Landscape:
                    return PageOrientationType.Landscape;
                default:
                    return PageOrientationType.Default;
            }
        }
        set
        {
            switch (value)
            {
                case PageOrientationType.Portrait:
                    _model.Orientation = PageOrientation.Portrait;
                    break;
                case PageOrientationType.Landscape:
                    _model.Orientation = PageOrientation.Landscape;
                    break;
                default:
                    _model.Orientation = PageOrientation.Default;
                    break;
            }
        }
    }

    public int? FirstPageNumber
    {
        get
        {
            return _model.FirstPageNumber;
        }
        set
        {
            if (value.HasValue && value.Value <= 0)
            {
                throw new CellsException("FirstPageNumber must be positive.");
            }

            _model.FirstPageNumber = value;
        }
    }

    public int? Scale
    {
        get
        {
            return _model.Scale;
        }
        set
        {
            if (value.HasValue && (value.Value < 10 || value.Value > 400))
            {
                throw new CellsException("Scale must be between 10 and 400.");
            }

            _model.Scale = value;
        }
    }

    public int? FitToPagesWide
    {
        get
        {
            return _model.FitToWidth;
        }
        set
        {
            if (value.HasValue && value.Value < 0)
            {
                throw new CellsException("FitToPagesWide must be zero or greater.");
            }

            _model.FitToWidth = value;
        }
    }

    public int? FitToPagesTall
    {
        get
        {
            return _model.FitToHeight;
        }
        set
        {
            if (value.HasValue && value.Value < 0)
            {
                throw new CellsException("FitToPagesTall must be zero or greater.");
            }

            _model.FitToHeight = value;
        }
    }

    public string PrintArea
    {
        get
        {
            return _model.PrintArea ?? string.Empty;
        }
        set
        {
            _model.PrintArea = NormalizeText(value);
        }
    }

    public string PrintTitleRows
    {
        get
        {
            return _model.PrintTitleRows ?? string.Empty;
        }
        set
        {
            _model.PrintTitleRows = NormalizeText(value);
        }
    }

    public string PrintTitleColumns
    {
        get
        {
            return _model.PrintTitleColumns ?? string.Empty;
        }
        set
        {
            _model.PrintTitleColumns = NormalizeText(value);
        }
    }

    public double LeftMargin
    {
        get
        {
            return ToCentimeters(_model.Margins.Left);
        }
        set
        {
            _model.Margins.Left = ValidateMargin(ToInches(value), nameof(LeftMargin));
        }
    }

    public double RightMargin
    {
        get
        {
            return ToCentimeters(_model.Margins.Right);
        }
        set
        {
            _model.Margins.Right = ValidateMargin(ToInches(value), nameof(RightMargin));
        }
    }

    public double TopMargin
    {
        get
        {
            return ToCentimeters(_model.Margins.Top);
        }
        set
        {
            _model.Margins.Top = ValidateMargin(ToInches(value), nameof(TopMargin));
        }
    }

    public double BottomMargin
    {
        get
        {
            return ToCentimeters(_model.Margins.Bottom);
        }
        set
        {
            _model.Margins.Bottom = ValidateMargin(ToInches(value), nameof(BottomMargin));
        }
    }

    public double HeaderMargin
    {
        get
        {
            return ToCentimeters(_model.Margins.Header);
        }
        set
        {
            _model.Margins.Header = ValidateMargin(ToInches(value), nameof(HeaderMargin));
        }
    }

    public double FooterMargin
    {
        get
        {
            return ToCentimeters(_model.Margins.Footer);
        }
        set
        {
            _model.Margins.Footer = ValidateMargin(ToInches(value), nameof(FooterMargin));
        }
    }

    public double LeftMarginInch
    {
        get
        {
            return _model.Margins.Left;
        }
        set
        {
            _model.Margins.Left = ValidateMargin(value, nameof(LeftMarginInch));
        }
    }

    public double RightMarginInch
    {
        get
        {
            return _model.Margins.Right;
        }
        set
        {
            _model.Margins.Right = ValidateMargin(value, nameof(RightMarginInch));
        }
    }

    public double TopMarginInch
    {
        get
        {
            return _model.Margins.Top;
        }
        set
        {
            _model.Margins.Top = ValidateMargin(value, nameof(TopMarginInch));
        }
    }

    public double BottomMarginInch
    {
        get
        {
            return _model.Margins.Bottom;
        }
        set
        {
            _model.Margins.Bottom = ValidateMargin(value, nameof(BottomMarginInch));
        }
    }

    public double HeaderMarginInch
    {
        get
        {
            return _model.Margins.Header;
        }
        set
        {
            _model.Margins.Header = ValidateMargin(value, nameof(HeaderMarginInch));
        }
    }

    public double FooterMarginInch
    {
        get
        {
            return _model.Margins.Footer;
        }
        set
        {
            _model.Margins.Footer = ValidateMargin(value, nameof(FooterMarginInch));
        }
    }

    public string LeftHeader
    {
        get
        {
            return _model.HeaderFooter.LeftHeader ?? string.Empty;
        }
        set
        {
            _model.HeaderFooter.LeftHeader = NormalizeText(value);
        }
    }

    public string CenterHeader
    {
        get
        {
            return _model.HeaderFooter.CenterHeader ?? string.Empty;
        }
        set
        {
            _model.HeaderFooter.CenterHeader = NormalizeText(value);
        }
    }

    public string RightHeader
    {
        get
        {
            return _model.HeaderFooter.RightHeader ?? string.Empty;
        }
        set
        {
            _model.HeaderFooter.RightHeader = NormalizeText(value);
        }
    }

    public string LeftFooter
    {
        get
        {
            return _model.HeaderFooter.LeftFooter ?? string.Empty;
        }
        set
        {
            _model.HeaderFooter.LeftFooter = NormalizeText(value);
        }
    }

    public string CenterFooter
    {
        get
        {
            return _model.HeaderFooter.CenterFooter ?? string.Empty;
        }
        set
        {
            _model.HeaderFooter.CenterFooter = NormalizeText(value);
        }
    }

    public string RightFooter
    {
        get
        {
            return _model.HeaderFooter.RightFooter ?? string.Empty;
        }
        set
        {
            _model.HeaderFooter.RightFooter = NormalizeText(value);
        }
    }

    public bool PrintGridlines
    {
        get
        {
            return _model.PrintOptions.GridLines;
        }
        set
        {
            _model.PrintOptions.GridLines = value;
        }
    }

    public bool PrintHeadings
    {
        get
        {
            return _model.PrintOptions.Headings;
        }
        set
        {
            _model.PrintOptions.Headings = value;
        }
    }

    public bool CenterHorizontally
    {
        get
        {
            return _model.PrintOptions.HorizontalCentered;
        }
        set
        {
            _model.PrintOptions.HorizontalCentered = value;
        }
    }

    public bool CenterVertically
    {
        get
        {
            return _model.PrintOptions.VerticalCentered;
        }
        set
        {
            _model.PrintOptions.VerticalCentered = value;
        }
    }

    public IReadOnlyList<int> HorizontalPageBreaks
    {
        get
        {
            return GetOrderedBreaks(_model.HorizontalPageBreaks);
        }
    }

    public IReadOnlyList<int> VerticalPageBreaks
    {
        get
        {
            return GetOrderedBreaks(_model.VerticalPageBreaks);
        }
    }

    public void AddHorizontalPageBreak(int rowIndex)
    {
        if (rowIndex < 0)
        {
            throw new CellsException("Horizontal page break row index must be non-negative.");
        }

        AddDistinct(_model.HorizontalPageBreaks, rowIndex);
    }

    public void AddVerticalPageBreak(int columnIndex)
    {
        if (columnIndex < 0)
        {
            throw new CellsException("Vertical page break column index must be non-negative.");
        }

        AddDistinct(_model.VerticalPageBreaks, columnIndex);
    }

    public void ClearHorizontalPageBreaks()
    {
        _model.HorizontalPageBreaks.Clear();
    }

    public void ClearVerticalPageBreaks()
    {
        _model.VerticalPageBreaks.Clear();
    }

    private static IReadOnlyList<int> GetOrderedBreaks(ICollection<int> breaks)
    {
        var orderedBreaks = new List<int>(breaks);
        orderedBreaks.Sort();
        return orderedBreaks;
    }

    private static void AddDistinct(ICollection<int> collection, int value)
    {
        if (!collection.Contains(value))
        {
            collection.Add(value);
        }
    }

    private static double ValidateMargin(double value, string propertyName)
    {
        if (value < 0d)
        {
            throw new CellsException($"{propertyName} must be zero or greater.");
        }

        return value;
    }

    private static double ToCentimeters(double inches)
    {
        return inches * CentimetersPerInch;
    }

    private static double ToInches(double centimeters)
    {
        return centimeters / CentimetersPerInch;
    }

    private static string? NormalizeText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return value!.Trim();
    }
}


