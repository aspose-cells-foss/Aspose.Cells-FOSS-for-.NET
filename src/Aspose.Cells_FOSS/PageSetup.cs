using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

/// <summary>
/// Represents worksheet print and page-layout settings.
/// </summary>
/// <example>
/// <code>
/// var workbook = new Workbook();
/// var pageSetup = workbook.Worksheets[0].PageSetup;
///
/// pageSetup.Orientation = PageOrientationType.Landscape;
/// pageSetup.LeftMargin = 1.5;
/// pageSetup.RightMargin = 1.5;
/// pageSetup.PrintTitleRows = "$1:$1";
/// pageSetup.AddHorizontalPageBreak(40);
/// </code>
/// </example>
public sealed class PageSetup
{
    private const double CentimetersPerInch = 2.54d;
    private readonly PageSetupModel _model;

    internal PageSetup(PageSetupModel model)
    {
        _model = model;
    }

    /// <summary>
    /// Gets or sets the paper size.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the page orientation.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the first printed page number.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the print scaling percentage.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the number of pages wide to fit when printing.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the number of pages tall to fit when printing.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the print area reference.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the repeating title rows reference.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the repeating title columns reference.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the left margin in centimeters.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the right margin in centimeters.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the top margin in centimeters.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the bottom margin in centimeters.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the header margin in centimeters.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the footer margin in centimeters.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the left margin in inches.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the right margin in inches.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the top margin in inches.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the bottom margin in inches.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the header margin in inches.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the footer margin in inches.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the left header text.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the center header text.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the right header text.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the left footer text.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the center footer text.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the right footer text.
    /// </summary>
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

    /// <summary>
    /// Gets or sets whether gridlines are printed.
    /// </summary>
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

    /// <summary>
    /// Gets or sets whether row and column headings are printed.
    /// </summary>
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

    /// <summary>
    /// Gets or sets whether content is centered horizontally on the page.
    /// </summary>
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

    /// <summary>
    /// Gets or sets whether content is centered vertically on the page.
    /// </summary>
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

    /// <summary>
    /// Gets the horizontal page breaks as sorted zero-based row indexes.
    /// </summary>
    public IReadOnlyList<int> HorizontalPageBreaks
    {
        get
        {
            return GetOrderedBreaks(_model.HorizontalPageBreaks);
        }
    }

    /// <summary>
    /// Gets the vertical page breaks as sorted zero-based column indexes.
    /// </summary>
    public IReadOnlyList<int> VerticalPageBreaks
    {
        get
        {
            return GetOrderedBreaks(_model.VerticalPageBreaks);
        }
    }

    /// <summary>
    /// Adds a horizontal page break at the specified zero-based row index.
    /// </summary>
    public void AddHorizontalPageBreak(int rowIndex)
    {
        if (rowIndex < 0)
        {
            throw new CellsException("Horizontal page break row index must be non-negative.");
        }

        AddDistinct(_model.HorizontalPageBreaks, rowIndex);
    }

    /// <summary>
    /// Adds a vertical page break at the specified zero-based column index.
    /// </summary>
    public void AddVerticalPageBreak(int columnIndex)
    {
        if (columnIndex < 0)
        {
            throw new CellsException("Vertical page break column index must be non-negative.");
        }

        AddDistinct(_model.VerticalPageBreaks, columnIndex);
    }

    /// <summary>
    /// Removes all horizontal page breaks.
    /// </summary>
    public void ClearHorizontalPageBreaks()
    {
        _model.HorizontalPageBreaks.Clear();
    }

    /// <summary>
    /// Removes all vertical page breaks.
    /// </summary>
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
