namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents the internal model for a worksheet comment (legacy note).
    /// Row and column values are zero-based.
    /// </summary>
    public sealed class CommentModel
    {
        /// <summary>
        /// Initializes a new instance with default field values.
        /// </summary>
        public CommentModel()
        {
            Author = string.Empty;
            Note = string.Empty;
            Width = 129;
            Height = 75;
        }

        /// <summary>
        /// Gets or sets the zero-based row index of the annotated cell.
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// Gets or sets the zero-based column index of the annotated cell.
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// Gets or sets the comment author name.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the plain-text note content.
        /// </summary>
        public string Note { get; set; }

        /// <summary>
        /// Gets or sets whether the comment box is always visible.
        /// </summary>
        public bool IsVisible { get; set; }

        /// <summary>
        /// Gets or sets the comment box width in pixels.
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Gets or sets the comment box height in pixels.
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Gets or sets the verbatim outer XML of the v:shape element from the VML drawing file.
        /// Null for programmatically-created comments; VML is generated at save time for those.
        /// </summary>
        internal string RawVmlShapeXml { get; set; }
    }
}
