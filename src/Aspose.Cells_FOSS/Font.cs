using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents font.
    /// </summary>
    public class Font
    {
        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        public string Name { get; set; } = "Calibri";
        /// <summary>
        /// Gets or sets the size.
        /// </summary>
        public double Size { get; set; } = 11d;
        /// <summary>
        /// Gets or sets a value indicating whether bold.
        /// </summary>
        public bool Bold { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether italic.
        /// </summary>
        public bool Italic { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether underline.
        /// </summary>
        public bool Underline { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether strike through.
        /// </summary>
        public bool StrikeThrough { get; set; }
        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        public Color Color { get; set; } = Color.Empty;

        /// <summary>
        /// Creates a copy of the current instance.
        /// </summary>
        /// <returns>The font.</returns>
        public Font Clone()
        {
            return new Font
            {
                Name = Name,
                Size = Size,
                Bold = Bold,
                Italic = Italic,
                Underline = Underline,
                StrikeThrough = StrikeThrough,
                Color = Color,
            };
        }
    }
}
