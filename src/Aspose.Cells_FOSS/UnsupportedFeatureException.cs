using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents an error that occurs during unsupported feature.
    /// </summary>
    public class UnsupportedFeatureException : CellsException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UnsupportedFeatureException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public UnsupportedFeatureException(string message) : base(message) { }
    }
}
