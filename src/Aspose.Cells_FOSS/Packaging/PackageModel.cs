using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Represents package model.
    /// </summary>
    internal sealed class PackageModel
    {
        /// <summary>
        /// Gets all package parts in the workbook package.
        /// </summary>
        public List<PackagePartDescriptor> Parts { get; } = new List<PackagePartDescriptor>();
        /// <summary>
        /// Gets all package relationships in the workbook package.
        /// </summary>
        public List<RelationshipDescriptor> Relationships { get; } = new List<RelationshipDescriptor>();
        /// <summary>
        /// Gets unsupported raw parts preserved during load/save.
        /// </summary>
        public Dictionary<string, byte[]> UnsupportedParts { get; } = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
    }
}

