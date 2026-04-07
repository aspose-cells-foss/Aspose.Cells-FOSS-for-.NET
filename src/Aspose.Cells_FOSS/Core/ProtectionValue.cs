using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class ProtectionValue
{
    public bool IsLocked { get; set; } = true;
    public bool IsHidden { get; set; }

    public ProtectionValue Clone()
    {
        return new ProtectionValue { IsLocked = IsLocked, IsHidden = IsHidden };
    }
}
