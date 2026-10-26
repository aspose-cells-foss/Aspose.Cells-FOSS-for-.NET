using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Numeric freeform paths preserved in drawing XML; guide formulas and arcs remain unsupported.</summary>
    internal static class ShapeCustomGeometry
    {
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static SKPath Build(XElement geometry, SKRect rect)
        {
            var paths = geometry.Element(A + "pathLst");
            if (paths == null) return null;
            var result = new SKPath();
            foreach (var path in paths.Elements(A + "path"))
            {
                double width, height;
                if (!TryNumber(path, "w", out width) || !TryNumber(path, "h", out height) || width <= 0d || height <= 0d)
                {
                    result.Dispose();
                    return null;
                }
                foreach (var command in path.Elements())
                {
                    var points = new List<SKPoint>();
                    foreach (var point in command.Elements(A + "pt"))
                    {
                        double x, y;
                        if (!TryNumber(point, "x", out x) || !TryNumber(point, "y", out y))
                        {
                            result.Dispose();
                            return null;
                        }
                        points.Add(new SKPoint(rect.Left + (float)(x / width * rect.Width), rect.Top + (float)(y / height * rect.Height)));
                    }
                    var name = command.Name.LocalName;
                    if (name == "moveTo" && points.Count == 1) result.MoveTo(points[0]);
                    else if (name == "lnTo" && points.Count == 1) result.LineTo(points[0]);
                    else if (name == "quadBezTo" && points.Count == 2) result.QuadTo(points[0], points[1]);
                    else if (name == "cubicBezTo" && points.Count == 3) result.CubicTo(points[0], points[1], points[2]);
                    else if (name == "close" && points.Count == 0) result.Close();
                    else
                    {
                        result.Dispose();
                        return null;
                    }
                }
            }
            return result;
        }

        private static bool TryNumber(XElement element, string name, out double value)
        {
            return double.TryParse((string)element.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                && !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}
