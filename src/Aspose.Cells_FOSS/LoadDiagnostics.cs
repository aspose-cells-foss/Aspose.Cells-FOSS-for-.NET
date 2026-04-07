using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public sealed class LoadDiagnostics
{
    private readonly List<LoadIssue> _issues = new List<LoadIssue>();

    public IReadOnlyList<LoadIssue> Issues
    {
        get
        {
            return _issues;
        }
    }

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
