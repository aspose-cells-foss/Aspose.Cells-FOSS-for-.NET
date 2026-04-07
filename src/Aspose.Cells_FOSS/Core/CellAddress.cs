using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public readonly struct CellAddress : IEquatable<CellAddress>
{
    public CellAddress(int rowIndex, int columnIndex)
    {
        if (rowIndex < 0) throw new ArgumentOutOfRangeException(nameof(rowIndex));
        if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
    }

    public int RowIndex { get; }
    public int ColumnIndex { get; }

    public static CellAddress Parse(string cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference)) throw new ArgumentException("Cell reference must be non-empty.", nameof(cellReference));
        var reference = cellReference.Trim();
        var index = 0;
        var column = 0;

        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            var letter = char.ToUpperInvariant(reference[index]);
            if (letter < 'A' || letter > 'Z') throw new ArgumentException($"Cell reference '{cellReference}' is invalid.", nameof(cellReference));
            column = (column * 26) + (letter - 'A' + 1);
            index++;
        }

        if (column == 0 || index == reference.Length) throw new ArgumentException($"Cell reference '{cellReference}' is invalid.", nameof(cellReference));

        var row = 0;
        while (index < reference.Length && char.IsDigit(reference[index]))
        {
            row = (row * 10) + (reference[index] - '0');
            index++;
        }

        if (index != reference.Length || row <= 0) throw new ArgumentException($"Cell reference '{cellReference}' is invalid.", nameof(cellReference));
        return new CellAddress(row - 1, column - 1);
    }

    public bool Equals(CellAddress other)
    {
        return RowIndex == other.RowIndex && ColumnIndex == other.ColumnIndex;
    }

    public override bool Equals(object? obj)
    {
        return obj is CellAddress other && Equals(other);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (RowIndex * 397) ^ ColumnIndex;
        }
    }

    public override string ToString()
    {
        return ColumnIndexToName(ColumnIndex) + (RowIndex + 1);
    }

    private static string ColumnIndexToName(int columnIndex)
    {
        var index = columnIndex + 1;
        var characters = new Stack<char>();
        while (index > 0)
        {
            index--;
            characters.Push((char)('A' + (index % 26)));
            index /= 26;
        }

        return new string(characters.ToArray());
    }
}