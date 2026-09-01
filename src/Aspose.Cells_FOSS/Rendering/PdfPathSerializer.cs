using System;
using System.Globalization;
using System.Text;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class PdfPathSerializer
    {
        public static string Serialize(SKPath path)
        {
            if (path == null)
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            var points = new SKPoint[4];
            var currentX = 0f;
            var currentY = 0f;
            var hasCurrentPoint = false;
            using (var iterator = path.CreateIterator(false))
            {
                while (true)
                {
                    var verb = iterator.Next(points);
                    if (verb == SKPathVerb.Done)
                    {
                        break;
                    }

                    if (verb == SKPathVerb.Move)
                    {
                        builder.Append(Format(points[0].X));
                        builder.Append(' ');
                        builder.Append(Format(points[0].Y));
                        builder.Append(" m\n");
                        currentX = points[0].X;
                        currentY = points[0].Y;
                        hasCurrentPoint = true;
                    }
                    else if (verb == SKPathVerb.Line)
                    {
                        builder.Append(Format(points[1].X));
                        builder.Append(' ');
                        builder.Append(Format(points[1].Y));
                        builder.Append(" l\n");
                        currentX = points[1].X;
                        currentY = points[1].Y;
                        hasCurrentPoint = true;
                    }
                    else if (verb == SKPathVerb.Quad)
                    {
                        var c1x = points[0].X + (points[1].X - points[0].X) * 2f / 3f;
                        var c1y = points[0].Y + (points[1].Y - points[0].Y) * 2f / 3f;
                        var c2x = points[2].X + (points[1].X - points[2].X) * 2f / 3f;
                        var c2y = points[2].Y + (points[1].Y - points[2].Y) * 2f / 3f;
                        AppendCubic(builder, ref currentX, ref currentY, ref hasCurrentPoint, c1x, c1y, c2x, c2y, points[2].X, points[2].Y);
                    }
                    else if (verb == SKPathVerb.Cubic)
                    {
                        AppendCubic(builder, ref currentX, ref currentY, ref hasCurrentPoint, points[1].X, points[1].Y, points[2].X, points[2].Y, points[3].X, points[3].Y);
                    }
                    else if (verb == SKPathVerb.Close)
                    {
                        builder.Append("h\n");
                        hasCurrentPoint = false;
                    }
                }
            }

            return builder.ToString();
        }

        private static void AppendCubic(StringBuilder builder, ref float currentX, ref float currentY, ref bool hasCurrentPoint, float c1x, float c1y, float c2x, float c2y, float x, float y)
        {
            if (hasCurrentPoint && SameCoordinate(c1x, currentX) && SameCoordinate(c1y, currentY))
            {
                builder.Append(Format(c2x));
                builder.Append(' ');
                builder.Append(Format(c2y));
                builder.Append(' ');
                builder.Append(Format(x));
                builder.Append(' ');
                builder.Append(Format(y));
                builder.Append(" v\n");
            }
            else if (SameCoordinate(c2x, x) && SameCoordinate(c2y, y))
            {
                builder.Append(Format(c1x));
                builder.Append(' ');
                builder.Append(Format(c1y));
                builder.Append(' ');
                builder.Append(Format(x));
                builder.Append(' ');
                builder.Append(Format(y));
                builder.Append(" y\n");
            }
            else
            {
                builder.Append(Format(c1x));
                builder.Append(' ');
                builder.Append(Format(c1y));
                builder.Append(' ');
                builder.Append(Format(c2x));
                builder.Append(' ');
                builder.Append(Format(c2y));
                builder.Append(' ');
                builder.Append(Format(x));
                builder.Append(' ');
                builder.Append(Format(y));
                builder.Append(" c\n");
            }

            currentX = x;
            currentY = y;
            hasCurrentPoint = true;
        }

        private static bool SameCoordinate(float left, float right)
        {
            return Math.Abs(left - right) <= 0.05f;
        }

        private static string Format(float value)
        {
            var rounded = (float)Math.Round(value, 1, MidpointRounding.AwayFromZero);
            if (rounded == 0f)
            {
                return "0";
            }

            return rounded.ToString("0.#", CultureInfo.InvariantCulture);
        }
    }
}
