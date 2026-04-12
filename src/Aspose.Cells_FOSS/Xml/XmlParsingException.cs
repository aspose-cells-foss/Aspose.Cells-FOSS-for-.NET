namespace Aspose.Cells_FOSS.Xml;

/// <summary>
/// Represents an error that occurs during xml parsing.
/// </summary>
public class XmlParsingException : Exception
{
    /// <summary>
    /// Initializes a new instance of the <see cref="XmlParsingException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public XmlParsingException(string message) : base(message) { }
}
