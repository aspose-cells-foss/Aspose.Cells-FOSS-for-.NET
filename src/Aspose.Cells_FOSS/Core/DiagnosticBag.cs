using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class DiagnosticBag
{
    private readonly List<DiagnosticEntry> _entries = new List<DiagnosticEntry>();

    public IReadOnlyList<DiagnosticEntry> Entries
    {
        get
        {
            return _entries;
        }
    }

    public void Add(DiagnosticEntry entry)
    {
        _entries.Add(entry);
    }
}
