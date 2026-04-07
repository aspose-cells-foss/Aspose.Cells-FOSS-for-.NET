using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

internal sealed class FormatConditionModel
{
    public FormatConditionType Type { get; set; }
    public OperatorType Operator { get; set; } = OperatorType.None;
    public string? Formula1 { get; set; }
    public string? Formula2 { get; set; }
    public string? TimePeriod { get; set; }
    public bool Duplicate { get; set; } = true;
    public bool Top { get; set; } = true;
    public bool Percent { get; set; }
    public int Rank { get; set; }
    public bool Above { get; set; } = true;
    public int StandardDeviation { get; set; }
    public int ColorScaleCount { get; set; } = 2;
    public ColorValue MinColor { get; set; }
    public ColorValue MidColor { get; set; }
    public ColorValue MaxColor { get; set; }
    public ColorValue BarColor { get; set; }
    public ColorValue NegativeBarColor { get; set; }
    public bool ShowBorder { get; set; }
    public string? Direction { get; set; }
    public string? BarLength { get; set; }
    public string? IconSetType { get; set; }
    public bool ReverseIcons { get; set; }
    public bool ShowIconOnly { get; set; }
    public int Priority { get; set; }
    public bool StopIfTrue { get; set; }
    public StyleValue Style { get; set; } = StyleValue.Default.Clone();
}
