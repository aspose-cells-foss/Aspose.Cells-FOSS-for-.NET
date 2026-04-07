using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public class InvalidFileFormatException : CellsException
{
    public InvalidFileFormatException(string message) : base(message) { }
    public InvalidFileFormatException(string message, Exception innerException) : base(message, innerException) { }
}
