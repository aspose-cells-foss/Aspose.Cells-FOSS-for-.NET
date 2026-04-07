using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class BordersValue
{
    public BorderSideValue Left { get; set; } = new BorderSideValue();
    public BorderSideValue Right { get; set; } = new BorderSideValue();
    public BorderSideValue Top { get; set; } = new BorderSideValue();
    public BorderSideValue Bottom { get; set; } = new BorderSideValue();
    public BorderSideValue Diagonal { get; set; } = new BorderSideValue();
    public bool DiagonalUp { get; set; }
    public bool DiagonalDown { get; set; }

    public BordersValue Clone()
    {
        return new BordersValue
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
