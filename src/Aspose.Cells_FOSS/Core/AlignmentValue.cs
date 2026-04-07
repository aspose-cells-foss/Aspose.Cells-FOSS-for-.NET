using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class AlignmentValue
{
    public HorizontalAlignment Horizontal { get; set; }
    public VerticalAlignment Vertical { get; set; } = VerticalAlignment.Bottom;
    public bool WrapText { get; set; }
    public int IndentLevel { get; set; }
    public int TextRotation { get; set; }
    public bool ShrinkToFit { get; set; }
    public int ReadingOrder { get; set; }
    public int RelativeIndent { get; set; }

    public AlignmentValue Clone()
    {
        return new AlignmentValue
        {
            Horizontal = Horizontal,
            Vertical = Vertical,
            WrapText = WrapText,
            IndentLevel = IndentLevel,
            TextRotation = TextRotation,
            ShrinkToFit = ShrinkToFit,
            ReadingOrder = ReadingOrder,
            RelativeIndent = RelativeIndent,
        };
    }
}
