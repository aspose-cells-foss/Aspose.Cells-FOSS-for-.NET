using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class BorderSideValue
{
    public BorderStyle Style { get; set; }
    public ColorValue Color { get; set; }
    public BorderSideValue Clone()
    {
        return new BorderSideValue { Style = Style, Color = Color };
    }
}
