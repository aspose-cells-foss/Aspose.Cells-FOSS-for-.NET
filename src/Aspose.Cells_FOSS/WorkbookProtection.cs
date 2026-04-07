using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class WorkbookProtection
{
    private readonly WorkbookProtectionModel _model;

    internal WorkbookProtection(WorkbookProtectionModel model)
    {
        _model = model;
    }

    public bool LockStructure
    {
        get
        {
            return _model.LockStructure;
        }
        set
        {
            _model.LockStructure = value;
        }
    }

    public bool LockWindows
    {
        get
        {
            return _model.LockWindows;
        }
        set
        {
            _model.LockWindows = value;
        }
    }

    public bool LockRevision
    {
        get
        {
            return _model.LockRevision;
        }
        set
        {
            _model.LockRevision = value;
        }
    }

    public string WorkbookPassword
    {
        get
        {
            return _model.WorkbookPassword;
        }
        set
        {
            _model.WorkbookPassword = value ?? string.Empty;
        }
    }

    public string RevisionsPassword
    {
        get
        {
            return _model.RevisionsPassword;
        }
        set
        {
            _model.RevisionsPassword = value ?? string.Empty;
        }
    }

    public bool IsProtected
    {
        get
        {
            return _model.HasStoredState();
        }
    }
}
