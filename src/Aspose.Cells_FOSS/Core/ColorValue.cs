using System.Linq;
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
        /// Initializes a new instance of the <see cref="ColorValue"/> class.
        /// </summary>
        /// <param name="a">The a.</param>
        /// <param name="r">The r.</param>
        /// <param name="g">The g.</param>
        /// <param name="b">The b.</param>
        public ColorValue(byte a, byte r, byte g, byte b)
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
        /// Determines whether the specified value is equal to the current instance.
        /// </summary>
        /// <param name="other">The other.</param>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool Equals(ColorValue other)
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
            return (obj is ColorValue)&& Equals(((ColorValue)obj));
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
    }
}
