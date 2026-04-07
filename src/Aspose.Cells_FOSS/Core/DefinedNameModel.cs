namespace Aspose.Cells_FOSS.Core;

public sealed class DefinedNameModel
{
    public string Name { get; set; } = string.Empty;
    public string Formula { get; set; } = string.Empty;
    public int? LocalSheetIndex { get; set; }
    public bool Hidden { get; set; }
    public string Comment { get; set; } = string.Empty;
}
