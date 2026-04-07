using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public class CellsException : Exception
{
    public CellsException(string message) : base(message) { }
    public CellsException(string message, Exception innerException) : base(message, innerException) { }
}
