using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents worksheet protection model.
    /// </summary>
    public sealed class WorksheetProtectionModel
    {
        /// <summary>
        /// Gets or sets a value indicating whether protected.
        /// </summary>
        public bool IsProtected { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether objects.
        /// </summary>
        public bool Objects { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether scenarios.
        /// </summary>
        public bool Scenarios { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether format cells.
        /// </summary>
        public bool FormatCells { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether format columns.
        /// </summary>
        public bool FormatColumns { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether format rows.
        /// </summary>
        public bool FormatRows { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether insert columns.
        /// </summary>
        public bool InsertColumns { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether insert rows.
        /// </summary>
        public bool InsertRows { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether insert hyperlinks.
        /// </summary>
        public bool InsertHyperlinks { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether delete columns.
        /// </summary>
        public bool DeleteColumns { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether delete rows.
        /// </summary>
        public bool DeleteRows { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether select locked cells.
        /// </summary>
        public bool SelectLockedCells { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether sort.
        /// </summary>
        public bool Sort { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether auto filter.
        /// </summary>
        public bool AutoFilter { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether pivot tables.
        /// </summary>
        public bool PivotTables { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether select unlocked cells.
        /// </summary>
        public bool SelectUnlockedCells { get; set; }
        /// <summary>
        /// Gets or sets the password hash.
        /// </summary>
        public string PasswordHash { get; set; }
        /// <summary>
        /// Gets or sets the algorithm name.
        /// </summary>
        public string AlgorithmName { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether h value.
        /// </summary>
        public string HashValue { get; set; }
        /// <summary>
        /// Gets or sets the salt value.
        /// </summary>
        public string SaltValue { get; set; }
        /// <summary>
        /// Gets or sets the spin count.
        /// </summary>
        public string SpinCount { get; set; }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            IsProtected = false;
            Objects = false;
            Scenarios = false;
            FormatCells = false;
            FormatColumns = false;
            FormatRows = false;
            InsertColumns = false;
            InsertRows = false;
            InsertHyperlinks = false;
            DeleteColumns = false;
            DeleteRows = false;
            SelectLockedCells = false;
            Sort = false;
            AutoFilter = false;
            PivotTables = false;
            SelectUnlockedCells = false;
            PasswordHash = null;
            AlgorithmName = null;
            HashValue = null;
            SaltValue = null;
            SpinCount = null;
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return IsProtected
                || Objects
                || Scenarios
                || FormatCells
                || FormatColumns
                || FormatRows
                || InsertColumns
                || InsertRows
                || InsertHyperlinks
                || DeleteColumns
                || DeleteRows
                || SelectLockedCells
                || Sort
                || AutoFilter
                || PivotTables
                || SelectUnlockedCells
                || !string.IsNullOrWhiteSpace(PasswordHash)
                || !string.IsNullOrWhiteSpace(AlgorithmName)
                || !string.IsNullOrWhiteSpace(HashValue)
                || !string.IsNullOrWhiteSpace(SaltValue)
                || !string.IsNullOrWhiteSpace(SpinCount);
        }
    }
}
