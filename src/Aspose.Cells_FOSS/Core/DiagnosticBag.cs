using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents diagnostic bag.
/// </summary>
public sealed class DiagnosticBag
{
    private readonly List<DiagnosticEntry> _entries = new List<DiagnosticEntry>();

    /// <summary>
    /// Gets the entries.
    /// </summary>
    public IReadOnlyList<DiagnosticEntry> Entries
    {
        get
        {
            return _entries;
        }
    }

    /// <summary>
    /// Adds the specified item.
    /// </summary>
    /// <param name="entry">The entry.</param>
    public void Add(DiagnosticEntry entry)
    {
        _entries.Add(entry);
    }
}
