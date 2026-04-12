using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents diagnostic entry.
/// </summary>
public sealed class DiagnosticEntry
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
    /// Gets or sets a value indicating whether repair applied.
    /// </summary>
    public bool RepairApplied { get; set; }
    /// <summary>
    /// Gets or sets a value indicating whether data loss risk.
    /// </summary>
    public bool DataLossRisk { get; set; }
}
