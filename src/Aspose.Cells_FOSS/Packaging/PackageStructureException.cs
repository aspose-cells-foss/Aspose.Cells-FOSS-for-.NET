using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Represents an error that occurs during package structure.
    /// </summary>
    internal class PackageStructureException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PackageStructureException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public PackageStructureException(string message) : base(message) { }
    }
}

