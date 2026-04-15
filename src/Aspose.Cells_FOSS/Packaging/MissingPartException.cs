using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Represents an error that occurs during missing part.
    /// </summary>
    public class MissingPartException : PackageStructureException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MissingPartException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public MissingPartException(string message) : base(message) { }
    }
}
