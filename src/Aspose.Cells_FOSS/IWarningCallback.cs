namespace Aspose.Cells_FOSS;

/// <summary>
/// Defines a callback that receives load warnings.
/// </summary>
public interface IWarningCallback
{
    /// <summary>
    /// Handles a warning raised while loading a workbook.
    /// </summary>
    /// <param name="warningInfo">The warning details.</param>
    void Warning(WarningInfo warningInfo);
}
