using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents an error that occurs during workbook save.
    /// </summary>
    public class WorkbookSaveException : CellsException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookSaveException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public WorkbookSaveException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookSaveException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The exception that caused the current exception.</param>
        public WorkbookSaveException(string message, Exception innerException) : base(message, innerException) { }
    }
}
