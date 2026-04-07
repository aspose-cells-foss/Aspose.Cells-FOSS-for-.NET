using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

internal sealed class ValidationModel
{
    public ValidationModel()
    {
        Areas = new List<CellArea>();
    }

    public List<CellArea> Areas { get; }
    public ValidationType Type { get; set; }
    public ValidationAlertType AlertStyle { get; set; } = ValidationAlertType.Stop;
    public OperatorType Operator { get; set; } = OperatorType.None;
    public string? Formula1 { get; set; }
    public string? Formula2 { get; set; }
    public bool IgnoreBlank { get; set; }
    public bool InCellDropDown { get; set; } = true;
    public string? InputTitle { get; set; }
    public string? InputMessage { get; set; }
    public string? ErrorTitle { get; set; }
    public string? ErrorMessage { get; set; }
    public bool ShowInput { get; set; }
    public bool ShowError { get; set; }
}