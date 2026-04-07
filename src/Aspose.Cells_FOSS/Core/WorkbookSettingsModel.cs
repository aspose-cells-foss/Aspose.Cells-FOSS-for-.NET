using System.Globalization;

namespace Aspose.Cells_FOSS.Core;

public sealed class WorkbookSettingsModel
{
    public DateSystem DateSystem { get; set; } = DateSystem.Windows1900;

    public CultureInfo DisplayCulture { get; set; } = CultureInfo.InvariantCulture;
}
