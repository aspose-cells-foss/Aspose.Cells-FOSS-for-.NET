using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents warning info.
    /// </summary>
    public sealed class WarningInfo
    {
        /// <summary>
        /// Gets or sets the code.
        /// </summary>
        public string Code { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the severity.
        /// </summary>
        public DiagnosticSeverity Severity { get; set; } = DiagnosticSeverity.Warning;
        /// <summary>
        /// Gets or sets the message.
        /// </summary>
        public string Message { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether data loss risk.
        /// </summary>
        public bool DataLossRisk { get; set; }
        /// <summary>
        /// Gets or sets the part uri.
        /// </summary>
        public string PartUri { get; set; }
        /// <summary>
        /// Gets or sets the sheet name.
        /// </summary>
        public string SheetName { get; set; }
        /// <summary>
        /// Gets or sets the cell ref.
        /// </summary>
        public string CellRef { get; set; }
        /// <summary>
        /// Gets or sets the row index.
        /// </summary>
        public int? RowIndex { get; set; }
    }
}
