using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents style value.
    /// </summary>
    public sealed class StyleValue
    {
        /// <summary>
        /// Gets the default.
        /// </summary>
        public static StyleValue Default
        {
            get
            {
                return new StyleValue();
            }
        }

        /// <summary>
        /// Performs font value.
        /// </summary>
        /// <returns>The font value font { get; set; } = new.</returns>
        public FontValue Font { get; set; } = new FontValue();
        /// <summary>
        /// Gets or sets the pattern.
        /// </summary>
        public FillPatternKind Pattern { get; set; }
        /// <summary>
        /// Gets or sets the foreground color.
        /// </summary>
        public ColorValue ForegroundColor { get; set; }
        /// <summary>
        /// Gets or sets the background color.
        /// </summary>
        public ColorValue BackgroundColor { get; set; }
        /// <summary>
        /// Performs borders value.
        /// </summary>
        /// <returns>The borders value borders { get; set; } = new.</returns>
        public BordersValue Borders { get; set; } = new BordersValue();
        /// <summary>
        /// Performs alignment value.
        /// </summary>
        /// <returns>The alignment value alignment { get; set; } = new.</returns>
        public AlignmentValue Alignment { get; set; } = new AlignmentValue();
        /// <summary>
        /// Performs protection value.
        /// </summary>
        /// <returns>The protection value protection { get; set; } = new.</returns>
        public ProtectionValue Protection { get; set; } = new ProtectionValue();
        /// <summary>
        /// Performs number format value.
        /// </summary>
        /// <returns>The number format value number format { get; set; } = new.</returns>
        public NumberFormatValue NumberFormat { get; set; } = new NumberFormatValue();

        /// <summary>
        /// Creates a copy of the current instance.
        /// </summary>
        /// <returns>The style value.</returns>
        public StyleValue Clone()
        {
            return new StyleValue
            {
                Font = Font.Clone(),
                Pattern = Pattern,
                ForegroundColor = ForegroundColor,
                BackgroundColor = BackgroundColor,
                Borders = Borders.Clone(),
                Alignment = Alignment.Clone(),
                Protection = Protection.Clone(),
                NumberFormat = NumberFormat.Clone(),
            };
        }
    }
}
