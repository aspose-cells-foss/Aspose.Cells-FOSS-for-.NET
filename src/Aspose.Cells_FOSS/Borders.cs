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
        /// Gets the left border settings.
        /// </summary>
        public Border Left { get; internal set; } = new Border();
        /// <summary>
        /// Gets the right border settings.
        /// </summary>
        public Border Right { get; internal set; } = new Border();
        /// <summary>
        /// Gets the top border settings.
        /// </summary>
        public Border Top { get; internal set; } = new Border();
        /// <summary>
        /// Gets the bottom border settings.
        /// </summary>
        public Border Bottom { get; internal set; } = new Border();
        /// <summary>
        /// Gets the diagonal border settings.
        /// </summary>
        public Border Diagonal { get; internal set; } = new Border();
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
        internal Borders Clone()
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
