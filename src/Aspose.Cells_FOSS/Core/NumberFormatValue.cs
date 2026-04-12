using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents number format value.
/// </summary>
public sealed class NumberFormatValue
{
    /// <summary>
    /// Gets or sets the number.
    /// </summary>
    public int Number { get; set; }
    /// <summary>
    /// Gets or sets the custom.
    /// </summary>
    public string? Custom { get; set; }

    /// <summary>
    /// Creates a copy of the current instance.
    /// </summary>
    /// <returns>The number format value.</returns>
    public NumberFormatValue Clone()
    {
        return new NumberFormatValue { Number = Number, Custom = Custom };
    }
}
