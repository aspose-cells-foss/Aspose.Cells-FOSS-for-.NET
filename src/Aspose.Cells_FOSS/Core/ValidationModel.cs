using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

internal sealed class ValidationModel
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ValidationModel"/> class.
    /// </summary>
    public ValidationModel()
    {
        Areas = new List<CellArea>();
    }

    /// <summary>
    /// Gets or sets the areas.
    /// </summary>
    public List<CellArea> Areas { get; }
    /// <summary>
    /// Gets or sets the type.
    /// </summary>
    public ValidationType Type { get; set; }
    /// <summary>
    /// Gets or sets the alert style.
    /// </summary>
    public ValidationAlertType AlertStyle { get; set; } = ValidationAlertType.Stop;
    /// <summary>
    /// Gets or sets the operator.
    /// </summary>
    public OperatorType Operator { get; set; } = OperatorType.None;
    /// <summary>
    /// Gets or sets the formula1.
    /// </summary>
    public string? Formula1 { get; set; }
    /// <summary>
    /// Gets or sets the formula2.
    /// </summary>
    public string? Formula2 { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether ignore blank.
    /// </summary>
    public bool IgnoreBlank { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether in cell drop down.
    /// </summary>
    public bool InCellDropDown { get; set; } = true;
    /// <summary>
    /// Gets or sets the input title.
    /// </summary>
    public string? InputTitle { get; set; }
    /// <summary>
    /// Gets or sets the input message.
    /// </summary>
    public string? InputMessage { get; set; }
    /// <summary>
    /// Gets or sets the error title.
    /// </summary>
    public string? ErrorTitle { get; set; }
    /// <summary>
    /// Gets or sets the error message.
    /// </summary>
    public string? ErrorMessage { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether show input.
    /// </summary>
    public bool ShowInput { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether show error.
    /// </summary>
    public bool ShowError { get; set; }
}
