using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents an error that occurs during invalid file format.
    /// </summary>
    public class InvalidFileFormatException : CellsException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InvalidFileFormatException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public InvalidFileFormatException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="InvalidFileFormatException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The exception that caused the current exception.</param>
        public InvalidFileFormatException(string message, Exception innerException) : base(message, innerException) { }
    }
}
