using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core;

/// <summary>
/// Represents workbook model.
/// </summary>
public sealed class WorkbookModel
{
    /// <summary>
    /// Initializes a new instance of the <see cref="WorkbookModel"/> class.
    /// </summary>
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

    /// <summary>
    /// Gets the worksheets.
    /// </summary>
    public List<WorksheetModel> Worksheets { get; }
    /// <summary>
    /// Gets the settings.
    /// </summary>
    public WorkbookSettingsModel Settings { get; }
    /// <summary>
    /// Gets the properties.
    /// </summary>
    public WorkbookPropertiesModel Properties { get; }
    /// <summary>
    /// Gets the document properties.
    /// </summary>
    public DocumentPropertiesModel DocumentProperties { get; }
    /// <summary>
    /// Gets the diagnostics.
    /// </summary>
    public DiagnosticBag Diagnostics { get; }
    /// <summary>
    /// Gets the styles.
    /// </summary>
    public StyleRepository Styles { get; }
    /// <summary>
    /// Gets or sets the shared strings.
    /// </summary>
    public SharedStringRepository SharedStrings { get; }
    /// <summary>
    /// Gets or sets the default style.
    /// </summary>
    public StyleValue DefaultStyle { get; set; }
    /// <summary>
    /// Gets or sets the active sheet index.
    /// </summary>
    public int ActiveSheetIndex { get; set; }
    /// <summary>
    /// Gets the defined names.
    /// </summary>
    public List<DefinedNameModel> DefinedNames { get; }
}
