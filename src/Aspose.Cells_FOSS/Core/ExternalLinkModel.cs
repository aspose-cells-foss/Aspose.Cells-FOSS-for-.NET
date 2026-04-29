namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Holds the raw XML for one external link part and its companion rels file.
    /// Both are preserved verbatim on round-trip.
    /// </summary>
    internal sealed class ExternalLinkModel
    {
        /// <summary>
        /// Gets or sets the verbatim XML content of xl/externalLinks/externalLink{N}.xml.
        /// </summary>
        public string RawXml { get; set; }

        /// <summary>
        /// Gets or sets the verbatim XML content of the external link's own .rels file,
        /// or null if that file was absent.
        /// </summary>
        public string RawRelsXml { get; set; }
    }
}
