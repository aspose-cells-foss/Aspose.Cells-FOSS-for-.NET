namespace Aspose.Cells_FOSS.Packaging;

/// <summary>
/// Provides packaging conventions operations.
/// </summary>
public static class PackagingConventions
{
    /// <summary>
    /// Gets the content types part.
    /// </summary>
    public const string ContentTypesPart = "/[Content_Types].xml";
    /// <summary>
    /// Gets the root relationships part.
    /// </summary>
    public const string RootRelationshipsPart = "/_rels/.rels";
    /// <summary>
    /// Gets the workbook part.
    /// </summary>
    public const string WorkbookPart = "/xl/workbook.xml";
    /// <summary>
    /// Gets the workbook relationships part.
    /// </summary>
    public const string WorkbookRelationshipsPart = "/xl/_rels/workbook.xml.rels";
    /// <summary>
    /// Gets the office document relationship.
    /// </summary>
    public const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    /// <summary>
    /// Gets the worksheet relationship.
    /// </summary>
    public const string WorksheetRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    /// <summary>
    /// Gets the styles relationship.
    /// </summary>
    public const string StylesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    /// <summary>
    /// Gets the shared strings relationship.
    /// </summary>
    public const string SharedStringsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
}
