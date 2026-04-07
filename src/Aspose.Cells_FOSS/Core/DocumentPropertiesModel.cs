using System;

namespace Aspose.Cells_FOSS.Core;

public sealed class DocumentPropertiesModel
{
    public DocumentPropertiesModel()
    {
        Core = new CoreDocumentPropertiesModel();
        Extended = new ExtendedDocumentPropertiesModel();
    }

    public CoreDocumentPropertiesModel Core { get; }
    public ExtendedDocumentPropertiesModel Extended { get; }

    public void CopyFrom(DocumentPropertiesModel source)
    {
        Core.CopyFrom(source.Core);
        Extended.CopyFrom(source.Extended);
    }

    public bool HasStoredState()
    {
        return Core.HasStoredState() || Extended.HasStoredState();
    }
}

public sealed class CoreDocumentPropertiesModel
{
    public string Title { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Creator { get; set; } = string.Empty;
    public string Keywords { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string LastModifiedBy { get; set; } = string.Empty;
    public string Revision { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string ContentStatus { get; set; } = string.Empty;
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }

    public void CopyFrom(CoreDocumentPropertiesModel source)
    {
        Title = source.Title;
        Subject = source.Subject;
        Creator = source.Creator;
        Keywords = source.Keywords;
        Description = source.Description;
        LastModifiedBy = source.LastModifiedBy;
        Revision = source.Revision;
        Category = source.Category;
        ContentStatus = source.ContentStatus;
        Created = source.Created;
        Modified = source.Modified;
    }

    public bool HasStoredState()
    {
        return !string.IsNullOrEmpty(Title)
            || !string.IsNullOrEmpty(Subject)
            || !string.IsNullOrEmpty(Creator)
            || !string.IsNullOrEmpty(Keywords)
            || !string.IsNullOrEmpty(Description)
            || !string.IsNullOrEmpty(LastModifiedBy)
            || !string.IsNullOrEmpty(Revision)
            || !string.IsNullOrEmpty(Category)
            || !string.IsNullOrEmpty(ContentStatus)
            || Created.HasValue
            || Modified.HasValue;
    }
}

public sealed class ExtendedDocumentPropertiesModel
{
    public string Application { get; set; } = string.Empty;
    public string AppVersion { get; set; } = string.Empty;
    public string Company { get; set; } = string.Empty;
    public string Manager { get; set; } = string.Empty;
    public int? DocSecurity { get; set; }
    public string HyperlinkBase { get; set; } = string.Empty;
    public bool? ScaleCrop { get; set; }
    public bool? LinksUpToDate { get; set; }
    public bool? SharedDoc { get; set; }

    public void CopyFrom(ExtendedDocumentPropertiesModel source)
    {
        Application = source.Application;
        AppVersion = source.AppVersion;
        Company = source.Company;
        Manager = source.Manager;
        DocSecurity = source.DocSecurity;
        HyperlinkBase = source.HyperlinkBase;
        ScaleCrop = source.ScaleCrop;
        LinksUpToDate = source.LinksUpToDate;
        SharedDoc = source.SharedDoc;
    }

    public bool HasStoredState()
    {
        return !string.IsNullOrEmpty(Application)
            || !string.IsNullOrEmpty(AppVersion)
            || !string.IsNullOrEmpty(Company)
            || !string.IsNullOrEmpty(Manager)
            || DocSecurity.HasValue
            || !string.IsNullOrEmpty(HyperlinkBase)
            || ScaleCrop.HasValue
            || LinksUpToDate.HasValue
            || SharedDoc.HasValue;
    }
}
