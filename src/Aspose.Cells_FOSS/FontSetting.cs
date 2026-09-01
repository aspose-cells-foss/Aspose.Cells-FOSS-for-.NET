using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a range of characters within a cell's rich text.
    /// </summary>
    public class FontSetting
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FontSetting"/> class.
        /// </summary>
        public FontSetting(int startIndex, int length)
        {
            if (startIndex < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(startIndex));
            }

            if (length < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(length));
            }

            StartIndex = startIndex;
            Length = length;
            Font = new Font();
        }

        /// <summary>
        /// Gets the zero-based start index of the run.
        /// </summary>
        public int StartIndex { get; private set; }

        /// <summary>
        /// Gets the run length.
        /// </summary>
        public int Length { get; private set; }

        /// <summary>
        /// Gets the font of the run.
        /// </summary>
        public Font Font { get; private set; }

        internal FontSetting Clone()
        {
            var clone = new FontSetting(StartIndex, Length);
            clone.Font = Font == null ? new Font() : Font.Clone();
            return clone;
        }

        internal RichTextRunValue ToCore()
        {
            return new RichTextRunValue
            {
                StartIndex = StartIndex,
                Length = Length,
                Font = new FontValue
                {
                    Name = Font == null ? "Calibri" : Font.Name,
                    Size = Font == null ? 11d : Font.Size,
                    Bold = Font != null && Font.IsBold,
                    Italic = Font != null && Font.IsItalic,
                    Underline = Font == null ? FontUnderlineType.None : Font.Underline,
                    StrikeThrough = Font != null && Font.IsStrikeout,
                    Color = Font == null ? default(ColorValue) : Font.Color.ToCore(),
                },
            };
        }

        internal static FontSetting FromCore(RichTextRunValue run)
        {
            var setting = new FontSetting(run.StartIndex, run.Length);
            if (run.Font != null)
            {
                setting.Font.Name = run.Font.Name;
                setting.Font.Size = run.Font.Size;
                setting.Font.IsBold = run.Font.Bold;
                setting.Font.IsItalic = run.Font.Italic;
                setting.Font.Underline = run.Font.Underline;
                setting.Font.IsStrikeout = run.Font.StrikeThrough;
                setting.Font.Color = Color.FromCore(run.Font.Color);
            }

            return setting;
        }
    }
}
