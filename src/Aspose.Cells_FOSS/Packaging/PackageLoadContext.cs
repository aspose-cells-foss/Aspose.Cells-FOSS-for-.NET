using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Packaging
{
    /// <summary>
    /// Represents package load context.
    /// </summary>
    public sealed class PackageLoadContext
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PackageLoadContext"/> class.
        /// </summary>
        /// <param name="workbook">The workbook.</param>
        /// <param name="package">The package.</param>
        public PackageLoadContext(object workbook, PackageModel package)
        {
            Workbook = workbook;
            Package = package;
        }

        /// <summary>
        /// Gets the workbook.
        /// </summary>
        public object Workbook { get; }
        /// <summary>
        /// Gets the package.
        /// </summary>
        public PackageModel Package { get; }
    }
}
