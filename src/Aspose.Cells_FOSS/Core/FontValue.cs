using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents font value.
/// </summary>
public sealed class FontValue
{
    /// <summary>
    /// Gets or sets the name.
    /// </summary>
    public string Name { get; set; } = "Calibri";
    /// <summary>
    /// Gets or sets the size.
    /// </summary>
    public double Size { get; set; } = 11d;
    /// <summary>
    /// Gets or sets a value indicating whether bold.
    /// </summary>
    public bool Bold { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether italic.
    /// </summary>
    public bool Italic { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether underline.
    /// </summary>
    public bool Underline { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether strike through.
    /// </summary>
    public bool StrikeThrough { get; set; }
    /// <summary>
    /// Gets or sets the color.
    /// </summary>
    public ColorValue Color { get; set; }

    /// <summary>
    /// Creates a copy of the current instance.
    /// </summary>
    /// <returns>The font value.</returns>
    public FontValue Clone()
    {
        return new FontValue
        {
            Name = Name,
            Size = Size,
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            StrikeThrough = StrikeThrough,
            Color = Color,
        };
    }
}
