using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

/// <summary>
/// Represents document properties.
/// </summary>
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

    /// <summary>
    /// Gets the core.
    /// </summary>
    public CoreDocumentProperties Core
    {
        get
        {
            return _core;
        }
    }

    /// <summary>
    /// Gets the extended.
    /// </summary>
    public ExtendedDocumentProperties Extended
    {
        get
        {
            return _extended;
        }
    }

    /// <summary>
    /// Gets or sets the title.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the subject.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the author.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the keywords.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the comments.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the category.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the company.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the manager.
    /// </summary>
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
