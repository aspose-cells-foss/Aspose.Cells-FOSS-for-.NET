using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents borders value.
    /// </summary>
    public sealed class BordersValue
    {
        /// <summary>
        /// Performs border side value.
        /// </summary>
        /// <returns>The border side value left { get; set; } = new.</returns>
        public BorderSideValue Left { get; set; } = new BorderSideValue();
        /// <summary>
        /// Performs border side value.
        /// </summary>
        /// <returns>The border side value right { get; set; } = new.</returns>
        public BorderSideValue Right { get; set; } = new BorderSideValue();
        /// <summary>
        /// Performs border side value.
        /// </summary>
        /// <returns>The border side value top { get; set; } = new.</returns>
        public BorderSideValue Top { get; set; } = new BorderSideValue();
        /// <summary>
        /// Performs border side value.
        /// </summary>
        /// <returns>The border side value bottom { get; set; } = new.</returns>
        public BorderSideValue Bottom { get; set; } = new BorderSideValue();
        /// <summary>
        /// Performs border side value.
        /// </summary>
        /// <returns>The border side value diagonal { get; set; } = new.</returns>
        public BorderSideValue Diagonal { get; set; } = new BorderSideValue();
        /// <summary>
        /// Gets or sets a value indicating whether diagonal up.
        /// </summary>
        public bool DiagonalUp { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether diagonal down.
        /// </summary>
        public bool DiagonalDown { get; set; }

        /// <summary>
        /// Creates a copy of the current instance.
        /// </summary>
        /// <returns>The borders value.</returns>
        public BordersValue Clone()
        {
            return new BordersValue
            {
                Left = Left.Clone(),
                Right = Right.Clone(),
                Top = Top.Clone(),
                Bottom = Bottom.Clone(),
                Diagonal = Diagonal.Clone(),
                DiagonalUp = DiagonalUp,
                DiagonalDown = DiagonalDown,
            };
        }
    }
}
