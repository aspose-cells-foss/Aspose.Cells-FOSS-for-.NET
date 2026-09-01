using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents one rich-text formatting run within a cell string.
    /// </summary>
    internal sealed class RichTextRunValue
    {
        /// <summary>
        /// Gets or sets the zero-based start index.
        /// </summary>
        public int StartIndex { get; set; }

        /// <summary>
        /// Gets or sets the character length.
        /// </summary>
        public int Length { get; set; }

        /// <summary>
        /// Gets or sets the font for the run.
        /// </summary>
        public FontValue Font { get; set; } = new FontValue();

        /// <summary>
        /// Creates a copy of the current run.
        /// </summary>
        public RichTextRunValue Clone()
        {
            return new RichTextRunValue
            {
                StartIndex = StartIndex,
                Length = Length,
                Font = Font == null ? new FontValue() : Font.Clone(),
            };
        }
    }
}
