namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents an image that is embedded inside a shape or group-shape raw XML (not a top-level
    /// worksheet picture). Preserved so that r:embed references in shape raw XML remain valid after save.
    /// </summary>
    internal sealed class ShapeImageModel
    {
        /// <summary>Gets or sets the original relationship ID (e.g. "rId2") from the drawing rels.</summary>
        public string OriginalRId { get; set; }

        /// <summary>Gets or sets the lowercase file extension without dot (e.g. "png", "jpeg").</summary>
        public string Extension { get; set; }

        /// <summary>Gets or sets the raw binary image bytes.</summary>
        public byte[] ImageData { get; set; }
    }
}
