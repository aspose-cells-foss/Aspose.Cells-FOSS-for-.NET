using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents color.
    /// </summary>
    public struct Color : IEquatable<Color>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Color"/> class.
        /// </summary>
        /// <param name="a">The a.</param>
        /// <param name="r">The r.</param>
        /// <param name="g">The g.</param>
        /// <param name="b">The b.</param>
        public Color(byte a, byte r, byte g, byte b)
        {
            A = a;
            R = r;
            G = g;
            B = b;
        }

        /// <summary>
        /// Gets the a.
        /// </summary>
        public byte A { get; }
        /// <summary>
        /// Gets the r.
        /// </summary>
        public byte R { get; }
        /// <summary>
        /// Gets the g.
        /// </summary>
        public byte G { get; }
        /// <summary>
        /// Gets the b.
        /// </summary>
        public byte B { get; }
        /// <summary>
        /// Gets the empty.
        /// </summary>
        public static Color Empty
        {
            get
            {
                return new Color(0, 0, 0, 0);
            }
        }

        /// <summary>
        /// Creates a color from ARGB components.
        /// </summary>
        /// <param name="a">The a.</param>
        /// <param name="r">The r.</param>
        /// <param name="g">The g.</param>
        /// <param name="b">The b.</param>
        /// <returns>The color.</returns>
        public static Color FromArgb(int a, int r, int g, int b)
        {
            return new Color((byte)a, (byte)r, (byte)g, (byte)b);
        }

        /// <summary>
        /// Determines whether the specified value is equal to the current instance.
        /// </summary>
        /// <param name="other">The other.</param>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool Equals(Color other)
        {
            return A == other.A && R == other.R && G == other.G && B == other.B;
        }

        /// <summary>
        /// Determines whether the specified value is equal to the current instance.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public override bool Equals(object obj)
        {
            return (obj is Color)&& Equals(((Color)obj));
        }

        /// <summary>
        /// Returns a hash code for the current instance.
        /// </summary>
        /// <returns>The int.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = A;
                hash = (hash * 397) ^ R;
                hash = (hash * 397) ^ G;
                hash = (hash * 397) ^ B;
                return hash;
            }
        }

        internal ColorValue ToCore()
        {
            return new ColorValue(A, R, G, B);
        }

        internal static Color FromCore(ColorValue value)
        {
            return new Color(value.A, value.R, value.G, value.B);
        }
    }
}
