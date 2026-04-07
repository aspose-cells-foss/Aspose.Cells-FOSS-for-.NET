using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public sealed class WarningInfo
{
    public string Code { get; set; } = string.Empty;
    public DiagnosticSeverity Severity { get; set; } = DiagnosticSeverity.Warning;
    public string Message { get; set; } = string.Empty;
    public bool DataLossRisk { get; set; }
    public string? PartUri { get; set; }
    public string? SheetName { get; set; }
    public string? CellRef { get; set; }
    public int? RowIndex { get; set; }
}
