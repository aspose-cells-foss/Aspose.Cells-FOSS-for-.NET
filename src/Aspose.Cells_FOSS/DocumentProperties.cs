using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class DocumentProperties
{
    private readonly DocumentPropertiesModel _model;
    private readonly CoreDocumentProperties _core;
    private readonly ExtendedDocumentProperties _extended;

    internal DocumentProperties(DocumentPropertiesModel model)
    {
        _model = model;
        _core = new CoreDocumentProperties(model.Core);
        _extended = new ExtendedDocumentProperties(model.Extended);
    }

    public CoreDocumentProperties Core
    {
        get
        {
            return _core;
        }
    }

    public ExtendedDocumentProperties Extended
    {
        get
        {
            return _extended;
        }
    }

    public string Title
    {
        get
        {
            return _model.Core.Title;
        }
        set
        {
            _model.Core.Title = value ?? string.Empty;
        }
    }

    public string Subject
    {
        get
        {
            return _model.Core.Subject;
        }
        set
        {
            _model.Core.Subject = value ?? string.Empty;
        }
    }

    public string Author
    {
        get
        {
            return _model.Core.Creator;
        }
        set
        {
            _model.Core.Creator = value ?? string.Empty;
        }
    }

    public string Keywords
    {
        get
        {
            return _model.Core.Keywords;
        }
        set
        {
            _model.Core.Keywords = value ?? string.Empty;
        }
    }

    public string Comments
    {
        get
        {
            return _model.Core.Description;
        }
        set
        {
            _model.Core.Description = value ?? string.Empty;
        }
    }

    public string Category
    {
        get
        {
            return _model.Core.Category;
        }
        set
        {
            _model.Core.Category = value ?? string.Empty;
        }
    }

    public string Company
    {
        get
        {
            return _model.Extended.Company;
        }
        set
        {
            _model.Extended.Company = value ?? string.Empty;
        }
    }

    public string Manager
    {
        get
        {
            return _model.Extended.Manager;
        }
        set
        {
            _model.Extended.Manager = value ?? string.Empty;
        }
    }
}

public sealed class CoreDocumentProperties
{
    private readonly CoreDocumentPropertiesModel _model;

    internal CoreDocumentProperties(CoreDocumentPropertiesModel model)
    {
        _model = model;
    }

    public string Title
    {
        get
        {
            return _model.Title;
        }
        set
        {
            _model.Title = value ?? string.Empty;
        }
    }

    public string Subject
    {
        get
        {
            return _model.Subject;
        }
        set
        {
            _model.Subject = value ?? string.Empty;
        }
    }

    public string Creator
    {
        get
        {
            return _model.Creator;
        }
        set
        {
            _model.Creator = value ?? string.Empty;
        }
    }

    public string Keywords
    {
        get
        {
            return _model.Keywords;
        }
        set
        {
            _model.Keywords = value ?? string.Empty;
        }
    }

    public string Description
    {
        get
        {
            return _model.Description;
        }
        set
        {
            _model.Description = value ?? string.Empty;
        }
    }

    public string LastModifiedBy
    {
        get
        {
            return _model.LastModifiedBy;
        }
        set
        {
            _model.LastModifiedBy = value ?? string.Empty;
        }
    }

    public string Revision
    {
        get
        {
            return _model.Revision;
        }
        set
        {
            _model.Revision = value ?? string.Empty;
        }
    }

    public string Category
    {
        get
        {
            return _model.Category;
        }
        set
        {
            _model.Category = value ?? string.Empty;
        }
    }

    public string ContentStatus
    {
        get
        {
            return _model.ContentStatus;
        }
        set
        {
            _model.ContentStatus = value ?? string.Empty;
        }
    }

    public DateTime? Created
    {
        get
        {
            return _model.Created;
        }
        set
        {
            _model.Created = value;
        }
    }

    public DateTime? Modified
    {
        get
        {
            return _model.Modified;
        }
        set
        {
            _model.Modified = value;
        }
    }
}

public sealed class ExtendedDocumentProperties
{
    private readonly ExtendedDocumentPropertiesModel _model;

    internal ExtendedDocumentProperties(ExtendedDocumentPropertiesModel model)
    {
        _model = model;
    }

    public string Application
    {
        get
        {
            return _model.Application;
        }
        set
        {
            _model.Application = value ?? string.Empty;
        }
    }

    public string AppVersion
    {
        get
        {
            return _model.AppVersion;
        }
        set
        {
            _model.AppVersion = value ?? string.Empty;
        }
    }

    public string Company
    {
        get
        {
            return _model.Company;
        }
        set
        {
            _model.Company = value ?? string.Empty;
        }
    }

    public string Manager
    {
        get
        {
            return _model.Manager;
        }
        set
        {
            _model.Manager = value ?? string.Empty;
        }
    }

    public int DocSecurity
    {
        get
        {
            return _model.DocSecurity ?? 0;
        }
        set
        {
            if (value < 0)
            {
                throw new CellsException("DocSecurity must be non-negative.");
            }

            _model.DocSecurity = value;
        }
    }

    public string HyperlinkBase
    {
        get
        {
            return _model.HyperlinkBase;
        }
        set
        {
            _model.HyperlinkBase = value ?? string.Empty;
        }
    }

    public bool ScaleCrop
    {
        get
        {
            return _model.ScaleCrop ?? false;
        }
        set
        {
            _model.ScaleCrop = value;
        }
    }

    public bool LinksUpToDate
    {
        get
        {
            return _model.LinksUpToDate ?? false;
        }
        set
        {
            _model.LinksUpToDate = value;
        }
    }

    public bool SharedDoc
    {
        get
        {
            return _model.SharedDoc ?? false;
        }
        set
        {
            _model.SharedDoc = value;
        }
    }
}
