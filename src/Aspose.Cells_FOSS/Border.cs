namespace Aspose.Cells_FOSS;

public class Border
{
    public BorderStyleType LineStyle { get; set; }
    public Color Color { get; set; } = Color.Empty;

    public Border Clone()
    {
        return new Border
        {
            LineStyle = LineStyle,
            Color = Color,
        };
    }
}
