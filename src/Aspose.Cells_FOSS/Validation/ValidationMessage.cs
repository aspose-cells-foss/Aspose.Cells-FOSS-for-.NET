namespace Aspose.Cells_FOSS.InternalValidation;

/// <summary>
/// Represents validation message.
/// </summary>
public sealed class ValidationMessage
{
    /// <summary>
    /// Gets or sets the code.
    /// </summary>
    public string Code { get; set; } = string.Empty;
    /// <summary>
    /// Gets or sets the severity.
    /// </summary>
    public ValidationMessageSeverity Severity { get; set; }
    /// <summary>
    /// Gets or sets the message.
    /// </summary>
    public string Message { get; set; } = string.Empty;
}
