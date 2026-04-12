using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

internal sealed class FormatConditionModel
{
    /// <summary>
    /// Gets or sets the type.
    /// </summary>
    public FormatConditionType Type { get; set; }
    /// <summary>
    /// Gets or sets the operator.
    /// </summary>
    public OperatorType Operator { get; set; } = OperatorType.None;
    /// <summary>
    /// Gets or sets the formula1.
    /// </summary>
    public string? Formula1 { get; set; }
    /// <summary>
    /// Gets or sets the formula2.
    /// </summary>
    public string? Formula2 { get; set; }
    /// <summary>
    /// Gets or sets the time period.
    /// </summary>
    public string? TimePeriod { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether duplicate.
    /// </summary>
    public bool Duplicate { get; set; } = true;
    /// <summary>
    /// Gets or sets a value indicating whether top.
    /// </summary>
    public bool Top { get; set; } = true;
    /// <summary>
    /// Gets or sets a value indicating whether percent.
    /// </summary>
    public bool Percent { get; set; }
    /// <summary>
    /// Gets or sets the rank.
    /// </summary>
    public int Rank { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether above.
    /// </summary>
    public bool Above { get; set; } = true;
    /// <summary>
    /// Gets or sets the standard deviation.
    /// </summary>
    public int StandardDeviation { get; set; }
    /// <summary>
    /// Gets or sets the color scale count.
    /// </summary>
    public int ColorScaleCount { get; set; } = 2;
    /// <summary>
    /// Gets or sets the min color.
    /// </summary>
    public ColorValue MinColor { get; set; }
    /// <summary>
    /// Gets or sets the mid color.
    /// </summary>
    public ColorValue MidColor { get; set; }
    /// <summary>
    /// Gets or sets the max color.
    /// </summary>
    public ColorValue MaxColor { get; set; }
    /// <summary>
    /// Gets or sets the bar color.
    /// </summary>
    public ColorValue BarColor { get; set; }
    /// <summary>
    /// Gets or sets the negative bar color.
    /// </summary>
    public ColorValue NegativeBarColor { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether show border.
    /// </summary>
    public bool ShowBorder { get; set; }
    /// <summary>
    /// Gets or sets the direction.
    /// </summary>
    public string? Direction { get; set; }
    /// <summary>
    /// Gets or sets the bar length.
    /// </summary>
    public string? BarLength { get; set; }
    /// <summary>
    /// Gets or sets the icon set type.
    /// </summary>
    public string? IconSetType { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether reverse icons.
    /// </summary>
    public bool ReverseIcons { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether show icon only.
    /// </summary>
    public bool ShowIconOnly { get; set; }
    /// <summary>
    /// Gets or sets the priority.
    /// </summary>
    public int Priority { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether stop if true.
    /// </summary>
    public bool StopIfTrue { get; set; }
    /// <summary>
    /// Performs style value.default.clone.
    /// </summary>
    /// <returns>The style value style { get; set; } =.</returns>
    public StyleValue Style { get; set; } = StyleValue.Default.Clone();
}
