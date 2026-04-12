namespace Aspose.Cells_FOSS;

/// <summary>
/// Represents border.
/// </summary>
public class Border
{
    /// <summary>
    /// Gets or sets the line style.
    /// </summary>
    public BorderStyleType LineStyle { get; set; }
    /// <summary>
    /// Gets or sets the color.
    /// </summary>
    public Color Color { get; set; } = Color.Empty;

    /// <summary>
    /// Creates a copy of the current instance.
    /// </summary>
    /// <returns>The border.</returns>
    public Border Clone()
    {
        return new Border
        {
            LineStyle = LineStyle,
            Color = Color,
        };
    }
}
