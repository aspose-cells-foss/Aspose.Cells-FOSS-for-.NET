using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

/// <summary>
/// Represents cell area.
/// </summary>
public readonly struct CellArea
{
    /// <summary>
    /// Initializes a new instance of the <see cref="CellArea"/> class.
    /// </summary>
    /// <param name="firstRow">The zero-based first row index.</param>
    /// <param name="firstColumn">The zero-based first column index.</param>
    /// <param name="totalRows">The total number of rows.</param>
    /// <param name="totalColumns">The total number of columns.</param>
    public CellArea(int firstRow, int firstColumn, int totalRows, int totalColumns)
    {
        if (firstRow < 0) throw new ArgumentOutOfRangeException(nameof(firstRow));
        if (firstColumn < 0) throw new ArgumentOutOfRangeException(nameof(firstColumn));
        if (totalRows <= 0) throw new ArgumentOutOfRangeException(nameof(totalRows));
        if (totalColumns <= 0) throw new ArgumentOutOfRangeException(nameof(totalColumns));

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
        if (endRow < startRow) throw new ArgumentOutOfRangeException(nameof(endRow));
        if (endColumn < startColumn) throw new ArgumentOutOfRangeException(nameof(endColumn));
        return new CellArea(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
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
}
