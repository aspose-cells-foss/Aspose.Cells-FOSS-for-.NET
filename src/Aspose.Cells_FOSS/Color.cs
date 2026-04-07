using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public readonly struct Color : IEquatable<Color>
{
    public Color(byte a, byte r, byte g, byte b)
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
    public static Color Empty
    {
        get
        {
            return new Color(0, 0, 0, 0);
        }
    }

    public static Color FromArgb(int a, int r, int g, int b)
    {
        return new Color((byte)a, (byte)r, (byte)g, (byte)b);
    }

    public bool Equals(Color other)
    {
        return A == other.A && R == other.R && G == other.G && B == other.B;
    }

    public override bool Equals(object? obj)
    {
        return obj is Color other && Equals(other);
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

    internal ColorValue ToCore()
    {
        return new ColorValue(A, R, G, B);
    }

    internal static Color FromCore(ColorValue value)
    {
        return new Color(value.A, value.R, value.G, value.B);
    }
}