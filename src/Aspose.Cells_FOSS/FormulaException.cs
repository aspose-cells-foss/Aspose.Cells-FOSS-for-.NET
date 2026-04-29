using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents an error that occurs during formula.
    /// </summary>
    public class FormulaException : CellsException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FormulaException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public FormulaException(string message) : base(message) { }
    }
}
