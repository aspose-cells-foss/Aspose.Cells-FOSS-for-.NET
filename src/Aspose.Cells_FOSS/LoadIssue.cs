using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents load issue.
    /// </summary>
    public sealed class LoadIssue
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="LoadIssue"/> class.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <param name="severity">The severity.</param>
        /// <param name="message">The message.</param>
        public LoadIssue(string code, DiagnosticSeverity severity, string message, bool repairApplied = false, bool dataLossRisk = false)
        {
            if (code == null) throw new ArgumentNullException(nameof(code));
            Code = code;
            Severity = severity;
            if (message == null) throw new ArgumentNullException(nameof(message));
            Message = message;
            RepairApplied = repairApplied;
            DataLossRisk = dataLossRisk;
        }

        /// <summary>
        /// Gets the code.
        /// </summary>
        public string Code { get; }
        /// <summary>
        /// Gets the severity.
        /// </summary>
        public DiagnosticSeverity Severity { get; }
        /// <summary>
        /// Gets the message.
        /// </summary>
        public string Message { get; }
        /// <summary>
        /// Gets a value indicating whether repair applied.
        /// </summary>
        public bool RepairApplied { get; }
        /// <summary>
        /// Gets or sets a value indicating whether data loss risk.
        /// </summary>
        public bool DataLossRisk { get; }
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
