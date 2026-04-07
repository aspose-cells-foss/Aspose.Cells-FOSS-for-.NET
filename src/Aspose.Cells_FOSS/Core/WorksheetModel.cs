using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class WorksheetModel
{
    public WorksheetModel(string name)
    {
        Name = name;
        Cells = new Dictionary<CellAddress, CellRecord>();
        Rows = new Dictionary<int, RowModel>();
        Columns = new List<ColumnRangeModel>();
        MergeRegions = new List<MergeRegion>();
        Hyperlinks = new List<HyperlinkModel>();
        Validations = new List<ValidationModel>();
        ConditionalFormattings = new List<ConditionalFormattingModel>();
        PageSetup = new PageSetupModel();
        View = new WorksheetViewModel();
        Protection = new WorksheetProtectionModel();
        AutoFilter = new AutoFilterModel();
    }

    public string Name { get; set; }
    public SheetVisibility Visibility { get; set; } = SheetVisibility.Visible;
    public Dictionary<CellAddress, CellRecord> Cells { get; }
    public Dictionary<int, RowModel> Rows { get; }
    public List<ColumnRangeModel> Columns { get; }
    public List<MergeRegion> MergeRegions { get; }
    public List<HyperlinkModel> Hyperlinks { get; }
    internal List<ValidationModel> Validations { get; }
    internal List<ConditionalFormattingModel> ConditionalFormattings { get; }
    public PageSetupModel PageSetup { get; }
    public WorksheetViewModel View { get; }
    public WorksheetProtectionModel Protection { get; }
    public AutoFilterModel AutoFilter { get; }
    public ColorValue? TabColor { get; set; }
}
