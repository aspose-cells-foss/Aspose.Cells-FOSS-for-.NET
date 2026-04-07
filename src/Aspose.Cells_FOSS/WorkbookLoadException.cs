using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public class WorkbookLoadException : CellsException
{
    public WorkbookLoadException(string message) : base(message) { }
    public WorkbookLoadException(string message, Exception innerException) : base(message, innerException) { }
}
