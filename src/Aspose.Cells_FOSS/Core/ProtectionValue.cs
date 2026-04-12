using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents protection value.
/// </summary>
public sealed class ProtectionValue
{
    /// <summary>
    /// Gets or sets a value indicating whether locked.
    /// </summary>
    public bool IsLocked { get; set; } = true;
    /// <summary>
    /// Gets or sets a value indicating whether hidden.
    /// </summary>
    public bool IsHidden { get; set; }

    /// <summary>
    /// Creates a copy of the current instance.
    /// </summary>
    /// <returns>The protection value.</returns>
    public ProtectionValue Clone()
    {
        return new ProtectionValue { IsLocked = IsLocked, IsHidden = IsHidden };
    }
}
