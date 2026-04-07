namespace Aspose.Cells_FOSS.Core;

public readonly struct MergeRegion
{
    public MergeRegion(int firstRow, int firstColumn, int totalRows, int totalColumns)
    {
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        TotalRows = totalRows;
        TotalColumns = totalColumns;
    }

    public int FirstRow { get; }
    public int FirstColumn { get; }
    public int TotalRows { get; }
    public int TotalColumns { get; }
}