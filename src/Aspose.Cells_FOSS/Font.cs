namespace Aspose.Cells_FOSS;

public class Font
{
    public string Name { get; set; } = "Calibri";
    public double Size { get; set; } = 11d;
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public bool StrikeThrough { get; set; }
    public Color Color { get; set; } = Color.Empty;

    public Font Clone()
    {
        return new Font
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
