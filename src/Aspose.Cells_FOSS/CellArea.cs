using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public readonly struct CellArea
{
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

    public int FirstRow { get; }
    public int FirstColumn { get; }
    public int TotalRows { get; }
    public int TotalColumns { get; }

    public static CellArea CreateCellArea(int startRow, int startColumn, int endRow, int endColumn)
    {
        if (endRow < startRow) throw new ArgumentOutOfRangeException(nameof(endRow));
        if (endColumn < startColumn) throw new ArgumentOutOfRangeException(nameof(endColumn));
        return new CellArea(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
    }

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
