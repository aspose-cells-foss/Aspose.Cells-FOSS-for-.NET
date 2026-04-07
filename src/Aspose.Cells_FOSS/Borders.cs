namespace Aspose.Cells_FOSS;

public class Borders
{
    public Border Left { get; set; } = new Border();
    public Border Right { get; set; } = new Border();
    public Border Top { get; set; } = new Border();
    public Border Bottom { get; set; } = new Border();
    public Border Diagonal { get; set; } = new Border();
    public bool DiagonalUp { get; set; }
    public bool DiagonalDown { get; set; }

    public Borders Clone()
    {
        return new Borders
        {
            Left = Left.Clone(),
            Right = Right.Clone(),
            Top = Top.Clone(),
            Bottom = Bottom.Clone(),
            Diagonal = Diagonal.Clone(),
            DiagonalUp = DiagonalUp,
            DiagonalDown = DiagonalDown,
        };
    }
}
