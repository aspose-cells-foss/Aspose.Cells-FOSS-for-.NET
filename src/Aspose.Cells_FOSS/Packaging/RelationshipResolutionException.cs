using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Represents an error that occurs during relationship resolution.
    /// </summary>
    internal class RelationshipResolutionException : PackageStructureException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RelationshipResolutionException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public RelationshipResolutionException(string message) : base(message) { }
    }
}

