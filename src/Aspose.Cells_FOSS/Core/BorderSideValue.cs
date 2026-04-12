using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents border side value.
/// </summary>
public sealed class BorderSideValue
{
    /// <summary>
    /// Gets or sets the style.
    /// </summary>
    public BorderStyle Style { get; set; }
    /// <summary>
    /// Gets or sets the color.
    /// </summary>
    public ColorValue Color { get; set; }
    /// <summary>
    /// Creates a copy of the current instance.
    /// </summary>
    /// <returns>The border side value.</returns>
    public BorderSideValue Clone()
    {
        return new BorderSideValue { Style = Style, Color = Color };
    }
}
