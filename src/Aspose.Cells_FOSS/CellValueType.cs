using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Specifies cell value type.
    /// </summary>
    public enum CellValueType
    {
        IsUnknown = 0,
        IsNull = 1,
        IsNumeric = 2,
        IsDateTime = 4,
        IsString = 8,
        IsBool = 16,
        IsError = 32
    }
}
