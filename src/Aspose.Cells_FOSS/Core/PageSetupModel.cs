using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class PageSetupModel
{
    public PageMarginsModel Margins { get; } = new PageMarginsModel();
    public PrintOptionsModel PrintOptions { get; } = new PrintOptionsModel();
    public HeaderFooterModel HeaderFooter { get; } = new HeaderFooterModel();
    public int PaperSize { get; set; }
    public PageOrientation Orientation { get; set; }
    public int? FirstPageNumber { get; set; }
    public int? Scale { get; set; }
    public int? FitToWidth { get; set; }
    public int? FitToHeight { get; set; }
    public string? PrintArea { get; set; }
    public string? PrintTitleRows { get; set; }
    public string? PrintTitleColumns { get; set; }
    public List<int> HorizontalPageBreaks { get; } = new List<int>();
    public List<int> VerticalPageBreaks { get; } = new List<int>();
}
