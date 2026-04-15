using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents borders.
    /// </summary>
    public class Borders
    {
        /// <summary>
        /// Performs border.
        /// </summary>
        /// <returns>The border left { get; set; } = new.</returns>
        public Border Left { get; set; } = new Border();
        /// <summary>
        /// Performs border.
        /// </summary>
        /// <returns>The border right { get; set; } = new.</returns>
        public Border Right { get; set; } = new Border();
        /// <summary>
        /// Performs border.
        /// </summary>
        /// <returns>The border top { get; set; } = new.</returns>
        public Border Top { get; set; } = new Border();
        /// <summary>
        /// Performs border.
        /// </summary>
        /// <returns>The border bottom { get; set; } = new.</returns>
        public Border Bottom { get; set; } = new Border();
        /// <summary>
        /// Performs border.
        /// </summary>
        /// <returns>The border diagonal { get; set; } = new.</returns>
        public Border Diagonal { get; set; } = new Border();
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
        /// <returns>The borders.</returns>
        public Borders Clone()
        {
            return new Borders
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
