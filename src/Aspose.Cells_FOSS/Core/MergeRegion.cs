using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents merge region.
    /// </summary>
    public struct MergeRegion
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MergeRegion"/> class.
        /// </summary>
        /// <param name="firstRow">The zero-based first row index.</param>
        /// <param name="firstColumn">The zero-based first column index.</param>
        /// <param name="totalRows">The total number of rows.</param>
        /// <param name="totalColumns">The total number of columns.</param>
        public MergeRegion(int firstRow, int firstColumn, int totalRows, int totalColumns)
        {
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            TotalRows = totalRows;
            TotalColumns = totalColumns;
        }

        /// <summary>
        /// Gets the first row.
        /// </summary>
        public int FirstRow { get; }
        /// <summary>
        /// Gets the first column.
        /// </summary>
        public int FirstColumn { get; }
        /// <summary>
        /// Gets the total rows.
        /// </summary>
        public int TotalRows { get; }
        /// <summary>
        /// Gets the total columns.
        /// </summary>
        public int TotalColumns { get; }
    }
}
