using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Represents package part descriptor.
    /// </summary>
    public sealed class PackagePartDescriptor
    {
        /// <summary>
        /// Gets or sets the part uri.
        /// </summary>
        public string PartUri { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the content type.
        /// </summary>
        public string ContentType { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the category.
        /// </summary>
        public string Category { get; set; } = string.Empty;
    }
}
