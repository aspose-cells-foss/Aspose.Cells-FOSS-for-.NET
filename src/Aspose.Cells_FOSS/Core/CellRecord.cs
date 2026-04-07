using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class CellRecord
{
    public object? Value { get; set; }
    public CellValueKind Kind { get; set; } = CellValueKind.Blank;
    public string? Formula { get; set; }
    public StyleValue Style { get; set; } = StyleValue.Default;
    public bool IsExplicitlyStored { get; set; }
}
