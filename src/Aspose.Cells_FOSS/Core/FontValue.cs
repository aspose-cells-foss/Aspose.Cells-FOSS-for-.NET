using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class FontValue
{
    public string Name { get; set; } = "Calibri";
    public double Size { get; set; } = 11d;
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public bool StrikeThrough { get; set; }
    public ColorValue Color { get; set; }

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
