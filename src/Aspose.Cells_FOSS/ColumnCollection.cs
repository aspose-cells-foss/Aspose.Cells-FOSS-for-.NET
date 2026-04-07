using System.Linq;

namespace Aspose.Cells_FOSS;

public sealed class ColumnCollection
{
    private readonly Worksheet _worksheet;

    internal ColumnCollection(Worksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public Column this[int index]
    {
        get
        {
            if (index < 0) throw new CellsException("Column index must be non-negative.");
            return new Column(_worksheet, index);
        }
    }
}
