using System.Collections.Generic;

namespace Aspose.Cells_FOSS;

public sealed class LoadOptions
{
    public LoadFormat LoadFormat { get; set; } = LoadFormat.Auto;

    public bool StrictMode { get; set; }

    public bool TryRepairPackage { get; set; } = true;

    public bool TryRepairXml { get; set; } = true;

    public bool PreserveUnsupportedParts { get; set; } = true;


    public IWarningCallback? WarningCallback { get; set; }
}
