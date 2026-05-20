using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents cell area.
    /// </summary>
    public struct CellArea : IComparable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CellArea"/> struct.
        /// </summary>
        internal CellArea(int firstRow, int firstColumn, int totalRows, int totalColumns)
        {
            if (firstRow < 0) throw new ArgumentOutOfRangeException(nameof(firstRow));
            if (firstColumn < 0) throw new ArgumentOutOfRangeException(nameof(firstColumn));
            if (totalRows <= 0) throw new ArgumentOutOfRangeException(nameof(totalRows));
            if (totalColumns <= 0) throw new ArgumentOutOfRangeException(nameof(totalColumns));

            StartRow = firstRow;
            StartColumn = firstColumn;
            EndRow = firstRow + totalRows - 1;
            EndColumn = firstColumn + totalColumns - 1;
        }

        /// <summary>
        /// Gets or sets the start row of this area.
        /// </summary>
        public int StartRow { get; set; }

        /// <summary>
        /// Gets or sets the end row of this area.
        /// </summary>
        public int EndRow { get; set; }

        /// <summary>
        /// Gets or sets the start column of this area.
        /// </summary>
        public int StartColumn { get; set; }

        /// <summary>
        /// Gets or sets the end column of this area.
        /// </summary>
        public int EndColumn { get; set; }

        internal int FirstRow { get { return StartRow; } }
        internal int FirstColumn { get { return StartColumn; } }
        internal int TotalRows { get { return EndRow - StartRow + 1; } }
        internal int TotalColumns { get { return EndColumn - StartColumn + 1; } }

        /// <summary>
        /// Creates the cell area.
        /// </summary>
        /// <param name="startRow">The start row.</param>
        /// <param name="startColumn">The start column.</param>
        /// <param name="endRow">The end row.</param>
        /// <param name="endColumn">The end column.</param>
        /// <returns>The cell area.</returns>
        public static CellArea CreateCellArea(int startRow, int startColumn, int endRow, int endColumn)
        {
            if (startRow < 0) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn < 0) throw new ArgumentOutOfRangeException(nameof(startColumn));
            if (endRow < startRow) throw new ArgumentOutOfRangeException(nameof(endRow));
            if (endColumn < startColumn) throw new ArgumentOutOfRangeException(nameof(endColumn));

            CellArea area = new CellArea();
            area.StartRow = startRow;
            area.StartColumn = startColumn;
            area.EndRow = endRow;
            area.EndColumn = endColumn;
            return area;
        }

        /// <summary>
        /// Creates the cell area.
        /// </summary>
        /// <param name="startCellName">The start cell reference.</param>
        /// <param name="endCellName">The end cell reference.</param>
        /// <returns>The cell area.</returns>
        public static CellArea CreateCellArea(string startCellName, string endCellName)
        {
            var start = CellAddress.Parse(startCellName);
            var end = CellAddress.Parse(endCellName);
            return CreateCellArea(
                Math.Min(start.RowIndex, end.RowIndex),
                Math.Min(start.ColumnIndex, end.ColumnIndex),
                Math.Max(start.RowIndex, end.RowIndex),
                Math.Max(start.ColumnIndex, end.ColumnIndex));
        }

        /// <summary>
        /// Compare two CellArea objects according to their top-left corner.
        /// </summary>
        /// <param name="obj">The object to compare.</param>
        /// <returns>The compare result.</returns>
        public int CompareTo(object obj)
        {
            if (obj == null)
            {
                return 1;
            }

            if (!(obj is CellArea))
            {
                throw new ArgumentException("Object must be a CellArea.", nameof(obj));
            }

            var other = (CellArea)obj;
            var rowComparison = StartRow.CompareTo(other.StartRow);
            if (rowComparison != 0)
            {
                return rowComparison;
            }

            var columnComparison = StartColumn.CompareTo(other.StartColumn);
            if (columnComparison != 0)
            {
                return columnComparison;
            }

            var endRowComparison = EndRow.CompareTo(other.EndRow);
            if (endRowComparison != 0)
            {
                return endRowComparison;
            }

            return EndColumn.CompareTo(other.EndColumn);
        }

        /// <summary>
        /// Returns a string represents the current cell area object.
        /// </summary>
        /// <returns>The text.</returns>
        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "R{0}C{1}:R{2}C{3}", StartRow, StartColumn, EndRow, EndColumn);
        }
    }
}
