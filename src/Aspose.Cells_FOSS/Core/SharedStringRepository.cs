using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents shared string repository.
/// </summary>
public sealed class SharedStringRepository
{
    private readonly Dictionary<string, int> _indices = new Dictionary<string, int>(StringComparer.Ordinal);
    private readonly List<string> _values = new List<string>();

    /// <summary>
    /// Gets the values.
    /// </summary>
    public IReadOnlyList<string> Values
    {
        get
        {
            return _values;
        }
    }

    /// <summary>
    /// Clears the current state.
    /// </summary>
    public void Clear()
    {
        _indices.Clear();
        _values.Clear();
    }

    /// <summary>
    /// Attempts to get value.
    /// </summary>
    /// <param name="index">The zero-based index.</param>
    /// <param name="value">The value.</param>
    /// <returns><see langword="true"/> if the operation succeeds; otherwise, <see langword="false"/>.</returns>
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

    /// <summary>
    /// Performs intern.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>The int.</returns>
    public int Intern(string value)
    {
        if (_indices.TryGetValue(value, out var index)) return index;
        index = _values.Count;
        _values.Add(value);
        _indices[value] = index;
        return index;
    }
}
