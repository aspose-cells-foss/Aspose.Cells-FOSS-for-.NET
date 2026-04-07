namespace Aspose.Cells_FOSS.InternalValidation;

public sealed class ValidationMessage
{
    public string Code { get; set; } = string.Empty;
    public ValidationMessageSeverity Severity { get; set; }
    public string Message { get; set; } = string.Empty;
}
