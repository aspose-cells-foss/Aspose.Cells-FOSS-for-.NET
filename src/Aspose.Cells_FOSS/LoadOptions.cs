using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Specifies how a workbook should be loaded.
    /// </summary>
    /// <example>
    /// <code>
    /// var workbook = new Workbook("input.xlsx", new LoadOptions
    /// {
    ///     StrictMode = false,
    ///     TryRepairPackage = true,
    ///     TryRepairXml = true,
    /// });
    /// </code>
    /// </example>
    public sealed class LoadOptions
    {
        /// <summary>
        /// Gets or sets the expected input format.
        /// </summary>
        public LoadFormat LoadFormat { get; set; } = LoadFormat.Auto;

        /// <summary>
        /// Gets or sets whether loading should reject ambiguous repairs.
        /// </summary>
        public bool StrictMode { get; set; }

        /// <summary>
        /// Gets or sets whether package-level repairs are allowed during load.
        /// </summary>
        public bool TryRepairPackage { get; set; } = true;

        /// <summary>
        /// Gets or sets whether XML-level repairs are allowed during load.
        /// </summary>
        public bool TryRepairXml { get; set; } = true;

        /// <summary>
        /// Gets or sets whether unsupported parts should be preserved when possible.
        /// </summary>
        public bool PreserveUnsupportedParts { get; set; } = true;

        /// <summary>
        /// Gets or sets a warning callback that receives recoverable-load diagnostics.
        /// </summary>
        public IWarningCallback WarningCallback { get; set; }
    }
}
