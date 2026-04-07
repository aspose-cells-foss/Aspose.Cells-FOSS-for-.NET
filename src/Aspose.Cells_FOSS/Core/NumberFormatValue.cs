using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class NumberFormatValue
{
    public int Number { get; set; }
    public string? Custom { get; set; }

    public NumberFormatValue Clone()
    {
        return new NumberFormatValue { Number = Number, Custom = Custom };
    }
}
