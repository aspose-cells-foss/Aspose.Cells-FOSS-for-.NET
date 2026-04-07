namespace Aspose.Cells_FOSS.Core;

public readonly struct ColorValue : IEquatable<ColorValue>
{
    public ColorValue(byte a, byte r, byte g, byte b)
    {
        A = a;
        R = r;
        G = g;
        B = b;
    }

    public byte A { get; }
    public byte R { get; }
    public byte G { get; }
    public byte B { get; }

    public bool Equals(ColorValue other)
    {
        return A == other.A && R == other.R && G == other.G && B == other.B;
    }

    public override bool Equals(object? obj)
    {
        return obj is ColorValue other && Equals(other);
    }

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