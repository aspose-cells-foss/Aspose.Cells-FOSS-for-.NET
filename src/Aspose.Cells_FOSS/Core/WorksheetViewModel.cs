namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents worksheet view model.
/// </summary>
public sealed class WorksheetViewModel
{
    /// <summary>
    /// Gets or sets a value indicating whether show grid lines.
    /// </summary>
    public bool ShowGridLines { get; set; } = true;
    /// <summary>
    /// Gets or sets a value indicating whether show row column headers.
    /// </summary>
    public bool ShowRowColumnHeaders { get; set; } = true;
    /// <summary>
    /// Gets or sets a value indicating whether show zeros.
    /// </summary>
    public bool ShowZeros { get; set; } = true;
    /// <summary>
    /// Gets or sets a value indicating whether right to left.
    /// </summary>
    public bool RightToLeft { get; set; }
    /// <summary>
    /// Gets or sets the zoom scale.
    /// </summary>
    public int ZoomScale { get; set; } = 100;
}
