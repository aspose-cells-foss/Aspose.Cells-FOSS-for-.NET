using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class DiagnosticEntry
{
    public string Code { get; set; } = string.Empty;
    public DiagnosticSeverity Severity { get; set; } = DiagnosticSeverity.Warning;
    public string Message { get; set; } = string.Empty;
    public bool RepairApplied { get; set; }
    public bool DataLossRisk { get; set; }
}
