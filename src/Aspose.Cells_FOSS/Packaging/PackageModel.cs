using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Represents package model.
    /// </summary>
    public sealed class PackageModel
    {
        /// <summary>
        /// Performs list<package part descriptor>.
        /// </summary>
        /// <returns>The list<package part descriptor> parts { get; } = new.</returns>
        public List<PackagePartDescriptor> Parts { get; } = new List<PackagePartDescriptor>();
        /// <summary>
        /// Performs list<relationship descriptor>.
        /// </summary>
        /// <returns>The list<relationship descriptor> relationships { get; } = new.</returns>
        public List<RelationshipDescriptor> Relationships { get; } = new List<RelationshipDescriptor>();
        /// <summary>
        /// Performs byte[]>.
        /// </summary>
        /// <param name="StringComparer.OrdinalIgnoreCase">The string comparer.ordinal ignore case.</param>
        /// <returns>The dictionary.</returns>
        public Dictionary<string, byte[]> UnsupportedParts { get; } = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
    }
}
