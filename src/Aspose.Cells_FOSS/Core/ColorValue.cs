using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents color value.
    /// </summary>
    public struct ColorValue : IEquatable<ColorValue>
    {
        /// <summary>
        /// Initializes an RGB color value.
        /// </summary>
        public ColorValue(byte a, byte r, byte g, byte b)
        {
            A = a;
            R = r;
            G = g;
            B = b;
            ThemeIndex = null;
            Tint = null;
            Indexed = null;
        }

        /// <summary>
        /// Initializes a theme-based color value.
        /// </summary>
        public ColorValue(int themeIndex, double? tint)
        {
            A = 0;
            R = 0;
            G = 0;
            B = 0;
            ThemeIndex = themeIndex;
            Tint = tint;
            Indexed = null;
        }

        /// <summary>
        /// Initializes an indexed color value.
        /// </summary>
        public ColorValue(int indexed)
        {
            A = 0;
            R = 0;
            G = 0;
            B = 0;
            ThemeIndex = null;
            Tint = null;
            Indexed = indexed;
        }

        /// <summary>Gets the alpha component (0 = transparent, 255 = opaque).</summary>
        public byte A { get; }
        /// <summary>Gets the red component.</summary>
        public byte R { get; }
        /// <summary>Gets the green component.</summary>
        public byte G { get; }
        /// <summary>Gets the blue component.</summary>
        public byte B { get; }
        /// <summary>Gets the theme color index, or null if not a theme color.</summary>
        public int? ThemeIndex { get; }
        /// <summary>Gets the tint/shade modifier applied to a theme color.</summary>
        public double? Tint { get; }
        /// <summary>Gets the indexed color value, or null if not an indexed color.</summary>
        public int? Indexed { get; }

        public bool Equals(ColorValue other)
        {
            return A == other.A && R == other.R && G == other.G && B == other.B
                && ThemeIndex == other.ThemeIndex && Tint == other.Tint && Indexed == other.Indexed;
        }

        public override bool Equals(object obj)
        {
            return (obj is ColorValue) && Equals(((ColorValue)obj));
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = A;
                hash = (hash * 397) ^ R;
                hash = (hash * 397) ^ G;
                hash = (hash * 397) ^ B;
                hash = (hash * 397) ^ (ThemeIndex ?? 0);
                hash = (hash * 397) ^ (Tint.HasValue ? Tint.Value.GetHashCode() : 0);
                hash = (hash * 397) ^ (Indexed ?? 0);
                return hash;
            }
        }
    }
}
