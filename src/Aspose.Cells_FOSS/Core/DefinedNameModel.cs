using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents defined name model.
    /// </summary>
    public sealed class DefinedNameModel
    {
        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the formula.
        /// </summary>
        public string Formula { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the local sheet index.
        /// </summary>
        public int? LocalSheetIndex { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether hidden.
        /// </summary>
        public bool Hidden { get; set; }
        /// <summary>
        /// Gets or sets the comment.
        /// </summary>
        public string Comment { get; set; } = string.Empty;
    }
}
