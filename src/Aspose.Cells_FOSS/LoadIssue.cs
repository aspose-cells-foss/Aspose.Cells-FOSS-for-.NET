using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public sealed class LoadIssue
{
    public LoadIssue(string code, DiagnosticSeverity severity, string message, bool repairApplied = false, bool dataLossRisk = false)
    {
        Code = code ?? throw new ArgumentNullException(nameof(code));
        Severity = severity;
        Message = message ?? throw new ArgumentNullException(nameof(message));
        RepairApplied = repairApplied;
        DataLossRisk = dataLossRisk;
    }

    public string Code { get; }
    public DiagnosticSeverity Severity { get; }
    public string Message { get; }
    public bool RepairApplied { get; }
    public bool DataLossRisk { get; }
    public string? PartUri { get; set; }
    public string? SheetName { get; set; }
    public string? CellRef { get; set; }
    public int? RowIndex { get; set; }
}
