using System.Linq;

namespace Aspose.Cells_FOSS;

public sealed class RowCollection
{
    private readonly Worksheet _worksheet;

    internal RowCollection(Worksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public Row this[int index]
    {
        get
        {
            if (index < 0) throw new CellsException("Row index must be non-negative.");
            return new Row(_worksheet, index);
        }
    }
}
