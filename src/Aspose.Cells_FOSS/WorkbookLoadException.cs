using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents an error that occurs during workbook load.
    /// </summary>
    public class WorkbookLoadException : CellsException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookLoadException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public WorkbookLoadException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookLoadException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The exception that caused the current exception.</param>
        public WorkbookLoadException(string message, Exception innerException) : base(message, innerException) { }
    }
}
