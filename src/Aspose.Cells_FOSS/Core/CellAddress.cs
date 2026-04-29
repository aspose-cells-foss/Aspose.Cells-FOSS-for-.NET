using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents cell address.
    /// </summary>
    public struct CellAddress : IEquatable<CellAddress>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CellAddress"/> class.
        /// </summary>
        /// <param name="rowIndex">The zero-based row index.</param>
        /// <param name="columnIndex">The zero-based column index.</param>
        public CellAddress(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0) throw new ArgumentOutOfRangeException(nameof(rowIndex));
            if (columnIndex < 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        /// <summary>
        /// Gets the row index.
        /// </summary>
        public int RowIndex { get; }
        /// <summary>
        /// Gets the column index.
        /// </summary>
        public int ColumnIndex { get; }

        /// <summary>
        /// Parses the specified value.
        /// </summary>
        /// <param name="cellReference">The cell reference.</param>
        /// <returns>The cell address.</returns>
        public static CellAddress Parse(string cellReference)
        {
            if (string.IsNullOrWhiteSpace(cellReference)) throw new ArgumentException("Cell reference must be non-empty.", nameof(cellReference));
            var reference = cellReference.Trim();
            var index = 0;
            var column = 0;

            // Convert the A1 column prefix into a zero-based index using base-26 letters.
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

        /// <summary>
        /// Determines whether the specified value is equal to the current instance.
        /// </summary>
        /// <param name="other">The other.</param>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool Equals(CellAddress other)
        {
            return RowIndex == other.RowIndex && ColumnIndex == other.ColumnIndex;
        }

        /// <summary>
        /// Determines whether the specified value is equal to the current instance.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public override bool Equals(object obj)
        {
            return (obj is CellAddress)&& Equals(((CellAddress)obj));
        }

        /// <summary>
        /// Returns a hash code for the current instance.
        /// </summary>
        /// <returns>The int.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                return (RowIndex * 397) ^ ColumnIndex;
            }
        }

        /// <summary>
        /// Returns the string representation of the current instance.
        /// </summary>
        /// <returns>The string.</returns>
        public override string ToString()
        {
            return ColumnIndexToName(ColumnIndex) + (RowIndex + 1);
        }

        private static string ColumnIndexToName(int columnIndex)
        {
            var index = columnIndex + 1;
            var characters = new Stack<char>();
            // Convert the zero-based column index back to Excel's repeated-letter form.
            while (index > 0)
            {
                index--;
                characters.Push((char)('A' + (index % 26)));
                index /= 26;
            }

            return new string(characters.ToArray());
        }
    }
}
