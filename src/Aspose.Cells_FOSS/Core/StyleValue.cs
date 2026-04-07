using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class StyleValue
{
    public static StyleValue Default
    {
        get
        {
            return new StyleValue();
        }
    }

    public FontValue Font { get; set; } = new FontValue();
    public FillPatternKind Pattern { get; set; }
    public ColorValue ForegroundColor { get; set; }
    public ColorValue BackgroundColor { get; set; }
    public BordersValue Borders { get; set; } = new BordersValue();
    public AlignmentValue Alignment { get; set; } = new AlignmentValue();
    public ProtectionValue Protection { get; set; } = new ProtectionValue();
    public NumberFormatValue NumberFormat { get; set; } = new NumberFormatValue();

    public StyleValue Clone()
    {
        return new StyleValue
        {
            Font = Font.Clone(),
            Pattern = Pattern,
            ForegroundColor = ForegroundColor,
            BackgroundColor = BackgroundColor,
            Borders = Borders.Clone(),
            Alignment = Alignment.Clone(),
            Protection = Protection.Clone(),
            NumberFormat = NumberFormat.Clone(),
        };
    }
}
