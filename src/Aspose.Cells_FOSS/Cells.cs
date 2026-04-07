using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

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

    public Cell this[string cellName]
    {
        get
        {
            try
            {
                return new Cell(_worksheet, CellAddress.Parse(cellName));
            }
            catch (ArgumentException exception)
            {
                throw new CellsException(exception.Message, exception);
            }
        }
    }

    public Cell this[int row, int column]
    {
        get
        {
            if (row < 0 || column < 0) throw new CellsException("Row and column indices must be non-negative.");
            return new Cell(_worksheet, new CellAddress(row, column));
        }
    }

    public RowCollection Rows
    {
        get
        {
            return _rows;
        }
    }

    public ColumnCollection Columns
    {
        get
        {
            return _columns;
        }
    }

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

    public void Merge(int firstRow, int firstColumn, int totalRows, int totalColumns)
    {
        if (firstRow < 0 || firstColumn < 0) throw new CellsException("Merge origin indices must be non-negative.");
        if (totalRows <= 0 || totalColumns <= 0) throw new CellsException("Merge range dimensions must be positive.");

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