using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public class WorkbookSaveException : CellsException
{
    public WorkbookSaveException(string message) : base(message) { }
    public WorkbookSaveException(string message, Exception innerException) : base(message, innerException) { }
}
