using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public sealed class SaveOptions
{
    public SaveFormat SaveFormat { get; set; } = SaveFormat.Xlsx;

    public bool UseSharedStrings { get; set; } = true;

    public bool ValidateBeforeSave { get; set; } = true;

    public bool CompactStyles { get; set; } = true;

    public bool PreserveRecoveryMetadata { get; set; }
}
