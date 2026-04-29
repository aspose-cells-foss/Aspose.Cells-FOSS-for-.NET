using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents an error that occurs during style.
    /// </summary>
    public class StyleException : CellsException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StyleException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public StyleException(string message) : base(message) { }
    }
}
