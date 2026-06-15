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
        /// Gets or sets the font size.
        /// Aspose-compatible property name.
        /// </summary>
        public double DoubleSize
        {
            get
            {
                return Size;
            }
            set
            {
                Size = value;
            }
        }
        /// <summary>
        /// Gets or sets the font underline type.
        /// </summary>
        public FontUnderlineType Underline { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether bold.
        /// Aspose-compatible property name.
        /// </summary>
        public bool IsBold { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether italic.
        /// Aspose-compatible property name.
        /// </summary>
        public bool IsItalic { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether single strikeout.
        /// Aspose-compatible property name.
        /// </summary>
        public bool IsStrikeout { get; set; }
        /// <summary>
        /// Gets or sets the color.
        /// </summary>
        public Color Color { get; set; } = Color.Empty;
        /// <summary>
        /// Gets or sets the color with a 32-bit ARGB value.
        /// Aspose-compatible property name.
        /// </summary>
        public int ArgbColor
        {
            get
            {
                return (Color.A << 24) | (Color.R << 16) | (Color.G << 8) | Color.B;
            }
            set
            {
                var a = (byte)((value >> 24) & 0xFF);
                var r = (byte)((value >> 16) & 0xFF);
                var g = (byte)((value >> 8) & 0xFF);
                var b = (byte)(value & 0xFF);
                Color = Color.FromArgb(a, r, g, b);
            }
        }

        /// <summary>
        /// Checks whether this font equals to another font instance.
        /// </summary>
        public bool Equals(Font font)
        {
            if (font == null)
            {
                return false;
            }

            return Name == font.Name
                && Size == font.Size
                && IsBold == font.IsBold
                && IsItalic == font.IsItalic
                && Underline == font.Underline
                && IsStrikeout == font.IsStrikeout
                && Color.Equals(font.Color);
        }

        /// <summary>
        /// Determines whether the specified object is equal to the current font instance.
        /// </summary>
        public override bool Equals(object obj)
        {
            return Equals(obj as Font);
        }

        /// <summary>
        /// Serves as a hash function for a font object.
        /// </summary>
        public override int GetHashCode()
        {
            unchecked
            {
                var hash = 17;
                hash = hash * 31 + (Name == null ? 0 : Name.GetHashCode());
                hash = hash * 31 + Size.GetHashCode();
                hash = hash * 31 + IsBold.GetHashCode();
                hash = hash * 31 + IsItalic.GetHashCode();
                hash = hash * 31 + Underline.GetHashCode();
                hash = hash * 31 + IsStrikeout.GetHashCode();
                hash = hash * 31 + Color.GetHashCode();
                return hash;
            }
        }

        /// <summary>
        /// Returns a string that represents the current font object.
        /// </summary>
        public override string ToString()
        {
            return string.Format(
                "Font(Name={0}, Size={1}, Bold={2}, Italic={3}, Underline={4}, Strikeout={5}, Argb={6})",
                Name,
                Size,
                IsBold,
                IsItalic,
                Underline,
                IsStrikeout,
                ArgbColor);
        }

        /// <summary>
        /// Creates a copy of the current instance.
        /// </summary>
        /// <returns>The font.</returns>
        internal Font Clone()
        {
            return new Font
            {
                Name = Name,
                Size = Size,
                IsBold = IsBold,
                IsItalic = IsItalic,
                Underline = Underline,
                IsStrikeout = IsStrikeout,
                Color = Color,
            };
        }
    }
}
