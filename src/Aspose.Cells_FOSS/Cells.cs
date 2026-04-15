using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Provides access to worksheet cells, rows, columns, and merged ranges.
    /// </summary>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// var cells = workbook.Worksheets[0].Cells;
    ///
    /// cells["A1"].PutValue("Total");
    /// cells[1, 1].PutValue(125.75);
    /// cells.Merge(0, 0, 1, 2);
    /// </code>
    /// </example>
    public class Cells
    {
        private readonly Worksheet _worksheet;
        private readonly RowCollection _rows;
        private readonly ColumnCollection _columns;

        internal Cells(Worksheet worksheet)
        {
            _worksheet = worksheet;
            _rows = new RowCollection(worksheet);
            _columns = new ColumnCollection(worksheet);
        }

        /// <summary>
        /// Gets a cell by A1 reference such as <c>A1</c> or <c>BC12</c>.
        /// </summary>
        public Cell this[string cellName]
        {
            get
            {
                try
                {
                    // The public API uses Aspose-style CellsException rather than exposing
                    // the lower-level argument validation coming from A1 parsing.
                    return new Cell(_worksheet, CellAddress.Parse(cellName));
                }
                catch (ArgumentException exception)
                {
                    throw new CellsException(exception.Message, exception);
                }
            }
        }

        /// <summary>
        /// Gets a cell by zero-based row and column index.
        /// </summary>
        public Cell this[int row, int column]
        {
            get
            {
                if (row < 0 || column < 0) throw new CellsException("Row and column indices must be non-negative.");
                return new Cell(_worksheet, new CellAddress(row, column));
            }
        }

        /// <summary>
        /// Gets row-level settings for the worksheet.
        /// </summary>
        public RowCollection Rows
        {
            get
            {
                return _rows;
            }
        }

        /// <summary>
        /// Gets column-level settings for the worksheet.
        /// </summary>
        public ColumnCollection Columns
        {
            get
            {
                return _columns;
            }
        }

        /// <summary>
        /// Gets the current merged-cell regions in worksheet order.
        /// </summary>
        public IReadOnlyList<CellArea> MergedCells
        {
            get
            {
                var mergedCells = new List<CellArea>(_worksheet.Model.MergeRegions.Count);
                foreach (var region in _worksheet.Model.MergeRegions)
                {
                    mergedCells.Add(new CellArea(region.FirstRow, region.FirstColumn, region.TotalRows, region.TotalColumns));
                }

                return mergedCells;
            }
        }

        /// <summary>
        /// Merges a rectangular cell region using zero-based coordinates.
        /// </summary>
        public void Merge(int firstRow, int firstColumn, int totalRows, int totalColumns)
        {
            if (firstRow < 0 || firstColumn < 0) throw new CellsException("Merge origin indices must be non-negative.");
            if (totalRows <= 0 || totalColumns <= 0) throw new CellsException("Merge range dimensions must be positive.");

            // Validate overlaps eagerly so the worksheet keeps a deterministic merge set
            // instead of deferring conflict resolution until save time.
            var candidate = new MergeRegion(firstRow, firstColumn, totalRows, totalColumns);
            foreach (var existing in _worksheet.Model.MergeRegions)
            {
                if (Overlaps(existing, candidate))
                {
                    throw new CellsException("Merge ranges must not overlap.");
                }
            }

            _worksheet.Model.MergeRegions.Add(candidate);
        }

        private static bool Overlaps(MergeRegion left, MergeRegion right)
        {
            var leftLastRow = left.FirstRow + left.TotalRows - 1;
            var leftLastColumn = left.FirstColumn + left.TotalColumns - 1;
            var rightLastRow = right.FirstRow + right.TotalRows - 1;
            var rightLastColumn = right.FirstColumn + right.TotalColumns - 1;

            return left.FirstRow <= rightLastRow
                && right.FirstRow <= leftLastRow
                && left.FirstColumn <= rightLastColumn
                && right.FirstColumn <= leftLastColumn;
        }
    }
}
