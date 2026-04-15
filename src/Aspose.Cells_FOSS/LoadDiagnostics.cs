using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents load diagnostics.
    /// </summary>
    public sealed class LoadDiagnostics
    {
        private readonly List<LoadIssue> _issues = new List<LoadIssue>();

        /// <summary>
        /// Gets a value indicating whether sues.
        /// </summary>
        public IReadOnlyList<LoadIssue> Issues
        {
            get
            {
                return _issues;
            }
        }

        /// <summary>
        /// Gets a value indicating whether repairs.
        /// </summary>
        public bool HasRepairs
        {
            get
            {
                foreach (var issue in _issues)
                {
                    if (issue.RepairApplied)
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether data loss risk.
        /// </summary>
        public bool HasDataLossRisk
        {
            get
            {
                foreach (var issue in _issues)
                {
                    if (issue.DataLossRisk)
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        internal void Add(LoadIssue issue)
        {
            _issues.Add(issue);
        }
    }
}
