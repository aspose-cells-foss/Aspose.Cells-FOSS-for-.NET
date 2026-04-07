using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

public sealed class WorkbookModel
{
    public WorkbookModel()
    {
        Settings = new WorkbookSettingsModel();
        Properties = new WorkbookPropertiesModel();
        DocumentProperties = new DocumentPropertiesModel();
        Diagnostics = new DiagnosticBag();
        Styles = new StyleRepository();
        SharedStrings = new SharedStringRepository();
        DefaultStyle = StyleValue.Default.Clone();
        Worksheets = new List<WorksheetModel>
        {
            new WorksheetModel("Sheet1"),
        };
        DefinedNames = new List<DefinedNameModel>();
        ActiveSheetIndex = 0;
    }

    public List<WorksheetModel> Worksheets { get; }
    public WorkbookSettingsModel Settings { get; }
    public WorkbookPropertiesModel Properties { get; }
    public DocumentPropertiesModel DocumentProperties { get; }
    public DiagnosticBag Diagnostics { get; }
    public StyleRepository Styles { get; }
    public SharedStringRepository SharedStrings { get; }
    public StyleValue DefaultStyle { get; set; }
    public int ActiveSheetIndex { get; set; }
    public List<DefinedNameModel> DefinedNames { get; }
}
