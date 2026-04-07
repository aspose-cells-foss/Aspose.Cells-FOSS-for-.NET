using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class WorkbookProperties
{
    private readonly WorkbookModel _workbookModel;
    private readonly WorkbookPropertiesModel _model;
    private readonly WorkbookProtection _protection;
    private readonly WorkbookView _view;
    private readonly CalculationProperties _calculation;

    internal WorkbookProperties(WorkbookModel workbookModel)
    {
        _workbookModel = workbookModel;
        _model = workbookModel.Properties;
        _protection = new WorkbookProtection(_model.Protection);
        _view = new WorkbookView(workbookModel);
        _calculation = new CalculationProperties(_model.Calculation);
    }

    public string CodeName
    {
        get
        {
            return _model.CodeName;
        }
        set
        {
            _model.CodeName = value ?? string.Empty;
        }
    }

    public string ShowObjects
    {
        get
        {
            return string.IsNullOrEmpty(_model.ShowObjects) ? "all" : _model.ShowObjects;
        }
        set
        {
            _model.ShowObjects = WorkbookPropertySupport.NormalizeShowObjects(value);
        }
    }

    public bool FilterPrivacy
    {
        get
        {
            return _model.FilterPrivacy;
        }
        set
        {
            _model.FilterPrivacy = value;
        }
    }

    public bool ShowBorderUnselectedTables
    {
        get
        {
            return _model.ShowBorderUnselectedTables;
        }
        set
        {
            _model.ShowBorderUnselectedTables = value;
        }
    }

    public bool ShowInkAnnotation
    {
        get
        {
            return _model.ShowInkAnnotation;
        }
        set
        {
            _model.ShowInkAnnotation = value;
        }
    }

    public bool BackupFile
    {
        get
        {
            return _model.BackupFile;
        }
        set
        {
            _model.BackupFile = value;
        }
    }

    public bool SaveExternalLinkValues
    {
        get
        {
            return _model.SaveExternalLinkValues;
        }
        set
        {
            _model.SaveExternalLinkValues = value;
        }
    }

    public string UpdateLinks
    {
        get
        {
            return string.IsNullOrEmpty(_model.UpdateLinks) ? "userSet" : _model.UpdateLinks;
        }
        set
        {
            _model.UpdateLinks = WorkbookPropertySupport.NormalizeUpdateLinks(value);
        }
    }

    public bool HidePivotFieldList
    {
        get
        {
            return _model.HidePivotFieldList;
        }
        set
        {
            _model.HidePivotFieldList = value;
        }
    }

    public int? DefaultThemeVersion
    {
        get
        {
            return _model.DefaultThemeVersion;
        }
        set
        {
            if (value.HasValue && value.Value < 0)
            {
                throw new CellsException("DefaultThemeVersion must be non-negative.");
            }

            _model.DefaultThemeVersion = value;
        }
    }

    public WorkbookProtection Protection
    {
        get
        {
            return _protection;
        }
    }

    public WorkbookView View
    {
        get
        {
            return _view;
        }
    }

    public CalculationProperties Calculation
    {
        get
        {
            return _calculation;
        }
    }
}
