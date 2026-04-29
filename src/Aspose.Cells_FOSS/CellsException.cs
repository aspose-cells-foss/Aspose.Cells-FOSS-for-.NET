using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents an error that occurs during cells.
    /// </summary>
    public class CellsException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CellsException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public CellsException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellsException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The exception that caused the current exception.</param>
        public CellsException(string message, Exception innerException) : base(message, innerException) { }
    }
}
