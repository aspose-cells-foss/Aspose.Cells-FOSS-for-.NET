using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class WorkbookSettings
{
    private readonly WorkbookSettingsModel _model;

    internal WorkbookSettings(WorkbookSettingsModel model)
    {
        _model = model;
    }

    public bool Date1904
    {
        get
        {
            return _model.DateSystem == Aspose.Cells_FOSS.Core.DateSystem.Mac1904;
        }
        set
        {
            _model.DateSystem = value ? Aspose.Cells_FOSS.Core.DateSystem.Mac1904 : Aspose.Cells_FOSS.Core.DateSystem.Windows1900;
        }
    }

    public CultureInfo Culture
    {
        get
        {
            return (CultureInfo)_model.DisplayCulture.Clone();
        }
        set
        {
            if (value is null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            _model.DisplayCulture = (CultureInfo)value.Clone();
        }
    }
}
