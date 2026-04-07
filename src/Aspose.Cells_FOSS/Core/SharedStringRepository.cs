using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class SharedStringRepository
{
    private readonly Dictionary<string, int> _indices = new Dictionary<string, int>(StringComparer.Ordinal);
    private readonly List<string> _values = new List<string>();

    public IReadOnlyList<string> Values
    {
        get
        {
            return _values;
        }
    }

    public void Clear()
    {
        _indices.Clear();
        _values.Clear();
    }

    public bool TryGetValue(int index, out string value)
    {
        if (index >= 0 && index < _values.Count)
        {
            value = _values[index];
            return true;
        }

        value = string.Empty;
        return false;
    }

    public int Intern(string value)
    {
        if (_indices.TryGetValue(value, out var index)) return index;
        index = _values.Count;
        _values.Add(value);
        _indices[value] = index;
        return index;
    }
}
