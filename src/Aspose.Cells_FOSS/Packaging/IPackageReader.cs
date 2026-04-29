using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Defines a reader for package models.
    /// </summary>
    public interface IPackageReader
    {
        /// <summary>
        /// Reads a package model from the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>The package model.</returns>
        PackageModel Read(Stream stream);
    }
}
