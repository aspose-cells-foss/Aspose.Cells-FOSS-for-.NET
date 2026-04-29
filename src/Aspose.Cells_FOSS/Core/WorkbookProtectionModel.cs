using System.IO;
using System.Collections.Generic;
using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents workbook protection model.
    /// </summary>
    public sealed class WorkbookProtectionModel
    {
        /// <summary>
        /// Gets or sets a value indicating whether lock structure.
        /// </summary>
        public bool LockStructure { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether lock windows.
        /// </summary>
        public bool LockWindows { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether lock revision.
        /// </summary>
        public bool LockRevision { get; set; }
        /// <summary>
        /// Gets or sets the workbook password.
        /// </summary>
        public string WorkbookPassword { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the revisions password.
        /// </summary>
        public string RevisionsPassword { get; set; } = string.Empty;

        /// <summary>
        /// Copies values from the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        public void CopyFrom(WorkbookProtectionModel source)
        {
            LockStructure = source.LockStructure;
            LockWindows = source.LockWindows;
            LockRevision = source.LockRevision;
            WorkbookPassword = source.WorkbookPassword;
            RevisionsPassword = source.RevisionsPassword;
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return LockStructure
                || LockWindows
                || LockRevision
                || !string.IsNullOrEmpty(WorkbookPassword)
                || !string.IsNullOrEmpty(RevisionsPassword);
        }
    }
}
