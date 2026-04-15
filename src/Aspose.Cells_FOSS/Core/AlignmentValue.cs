using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents alignment value.
    /// </summary>
    public sealed class AlignmentValue
    {
        /// <summary>
        /// Gets or sets the horizontal.
        /// </summary>
        public HorizontalAlignment Horizontal { get; set; }
        /// <summary>
        /// Gets or sets the vertical.
        /// </summary>
        public VerticalAlignment Vertical { get; set; } = VerticalAlignment.Bottom;
        /// <summary>
        /// Gets or sets a value indicating whether wrap text.
        /// </summary>
        public bool WrapText { get; set; }
        /// <summary>
        /// Gets or sets the indent level.
        /// </summary>
        public int IndentLevel { get; set; }
        /// <summary>
        /// Gets or sets the text rotation.
        /// </summary>
        public int TextRotation { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether shrink to fit.
        /// </summary>
        public bool ShrinkToFit { get; set; }
        /// <summary>
        /// Gets or sets the reading order.
        /// </summary>
        public int ReadingOrder { get; set; }
        /// <summary>
        /// Gets or sets the relative indent.
        /// </summary>
        public int RelativeIndent { get; set; }

        /// <summary>
        /// Creates a copy of the current instance.
        /// </summary>
        /// <returns>The alignment value.</returns>
        public AlignmentValue Clone()
        {
            return new AlignmentValue
            {
                Horizontal = Horizontal,
                Vertical = Vertical,
                WrapText = WrapText,
                IndentLevel = IndentLevel,
                TextRotation = TextRotation,
                ShrinkToFit = ShrinkToFit,
                ReadingOrder = ReadingOrder,
                RelativeIndent = RelativeIndent,
            };
        }
    }
}
