using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class StyleRepository
{
    public StyleValue Normalize(StyleValue style)
    {
        return style.Clone();
    }
}
