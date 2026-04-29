using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Represents relationship descriptor.
    /// </summary>
    public sealed class RelationshipDescriptor
    {
        /// <summary>
        /// Gets or sets the id.
        /// </summary>
        public string Id { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the target.
        /// </summary>
        public string Target { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether external.
        /// </summary>
        public bool IsExternal { get; set; }
    }
}
