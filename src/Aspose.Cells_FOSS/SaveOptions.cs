using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Specifies how a workbook should be saved.
    /// </summary>
    /// <example>
    /// <code>
    /// workbook.Save("output.xlsx", new SaveOptions
    /// {
    ///     UseSharedStrings = true,
    ///     ValidateBeforeSave = true,
    ///     CompactStyles = true,
    /// });
    /// </code>
    /// </example>
    public sealed class SaveOptions
    {
        /// <summary>
        /// Gets or sets the output file format.
        /// </summary>
        public SaveFormat SaveFormat { get; set; } = SaveFormat.Xlsx;

        /// <summary>
        /// Gets or sets whether shared strings should be used for string cells.
        /// </summary>
        public bool UseSharedStrings { get; set; } = true;

        /// <summary>
        /// Gets or sets whether the workbook should be validated before save.
        /// </summary>
        public bool ValidateBeforeSave { get; set; } = true;

        /// <summary>
        /// Gets or sets whether equivalent styles should be compacted during save.
        /// </summary>
        public bool CompactStyles { get; set; } = true;

        /// <summary>
        /// Gets or sets whether recovery metadata should be preserved in the saved workbook.
        /// </summary>
        public bool PreserveRecoveryMetadata { get; set; }
    }
}
