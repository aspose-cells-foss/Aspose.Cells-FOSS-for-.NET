using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

internal sealed class ConditionalFormattingModel
{
    public ConditionalFormattingModel()
    {
        Areas = new List<CellArea>();
        Conditions = new List<FormatConditionModel>();
    }

    public List<CellArea> Areas { get; }
    public List<FormatConditionModel> Conditions { get; }
}
