using System;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartTextStyleResolver
    {
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static bool TryResolveFont(XElement defaultRunProperties, out FontValue font)
        {
            font = null;
            if (defaultRunProperties == null)
            {
                return false;
            }

            var resolved = new FontValue();
            var found = false;

            double sizePt;
            if (TryResolveSize(defaultRunProperties, out sizePt))
            {
                resolved.Size = sizePt;
                found = true;
            }

            bool bold;
            if (TryResolveBooleanAttr(defaultRunProperties, "b", out bold))
            {
                resolved.Bold = bold;
                found = true;
            }

            bool italic;
            if (TryResolveBooleanAttr(defaultRunProperties, "i", out italic))
            {
                resolved.Italic = italic;
                found = true;
            }

            string typeface;
            if (TryResolveTypeface(defaultRunProperties, out typeface))
            {
                resolved.Name = typeface;
                found = true;
            }

            if (!found)
            {
                return false;
            }

            font = resolved;
            return true;
        }

        public static bool TryResolveTextColor(XElement defaultRunProperties, RenderColor colors, ChartModel chartModel, out SKColor color)
        {
            color = SKColors.Transparent;
            if (defaultRunProperties == null)
            {
                return false;
            }

            var solidFill = defaultRunProperties.Element(A + "solidFill");
            var fillSource = ChartXmlParser.FillChild(solidFill);
            if (ChartXmlParser.TryDrawingColor(fillSource, colors, out color))
            {
                return true;
            }

            var gradientFill = defaultRunProperties.Element(A + "gradFill");
            if (TryResolveGradientTextColor(gradientFill, colors, out color))
            {
                return true;
            }

            var blipFill = defaultRunProperties.Element(A + "blipFill");
            if (blipFill == null || chartModel == null)
            {
                return false;
            }

            var blip = blipFill.Element(A + "blip");
            var relationshipId = blip != null ? AttrOf(blip, R + "embed") : null;
            if (string.IsNullOrEmpty(relationshipId))
            {
                relationshipId = blip != null ? AttrOf(blip, "embed") : null;
            }

            if (string.IsNullOrEmpty(relationshipId))
            {
                return false;
            }

            ChartCompanionFile companion = null;
            for (var i = 0; i < chartModel.CompanionFiles.Count; i++)
            {
                var candidate = chartModel.CompanionFiles[i];
                if (candidate != null
                    && string.Equals(candidate.RelationshipId, relationshipId, StringComparison.Ordinal)
                    && candidate.BinaryContent != null
                    && candidate.BinaryContent.Length > 0)
                {
                    companion = candidate;
                    break;
                }
            }

            if (companion == null)
            {
                return false;
            }

            return TryResolveRepresentativeImageColor(companion.BinaryContent, out color);
        }

        private static bool TryResolveGradientTextColor(XElement gradientFill, RenderColor colors, out SKColor color)
        {
            color = SKColors.Transparent;
            if (gradientFill == null || colors == null)
            {
                return false;
            }

            var gradientStops = gradientFill.Element(A + "gsLst");
            if (gradientStops == null)
            {
                return false;
            }

            var resolvedStops = new System.Collections.Generic.List<System.Tuple<int, SKColor>>();
            foreach (var stop in gradientStops.Elements(A + "gs"))
            {
                int position;
                if (!int.TryParse(AttrOf(stop, "pos"), NumberStyles.Integer, CultureInfo.InvariantCulture, out position))
                {
                    continue;
                }

                var source = stop.Element(A + "srgbClr") ?? stop.Element(A + "schemeClr");
                SKColor stopColor;
                if (ChartXmlParser.TryDrawingColor(source, colors, out stopColor))
                {
                    resolvedStops.Add(new System.Tuple<int, SKColor>(position, stopColor));
                }
            }

            if (resolvedStops.Count == 0)
            {
                return false;
            }

            if (resolvedStops.Count == 1)
            {
                color = resolvedStops[0].Item2;
                return true;
            }

            double totalWeight = 0d;
            double sumR = 0d;
            double sumG = 0d;
            double sumB = 0d;
            double sumA = 0d;
            for (var i = 0; i < resolvedStops.Count; i++)
            {
                var current = resolvedStops[i];
                var previousPos = i == 0 ? 0 : resolvedStops[i - 1].Item1;
                var nextPos = i == resolvedStops.Count - 1 ? 100000 : resolvedStops[i + 1].Item1;
                var weight = Math.Max(1d, (nextPos - previousPos) * 0.5d);
                totalWeight += weight;
                sumR += current.Item2.Red * weight;
                sumG += current.Item2.Green * weight;
                sumB += current.Item2.Blue * weight;
                sumA += current.Item2.Alpha * weight;
            }

            if (totalWeight <= 0d)
            {
                return false;
            }

            color = new SKColor(
                (byte)Math.Round(sumR / totalWeight),
                (byte)Math.Round(sumG / totalWeight),
                (byte)Math.Round(sumB / totalWeight),
                (byte)Math.Round(sumA / totalWeight));
            return true;
        }

        private static bool TryResolveRepresentativeImageColor(byte[] imageBytes, out SKColor color)
        {
            color = SKColors.Transparent;
            if (imageBytes == null || imageBytes.Length == 0)
            {
                return false;
            }

            using (var bitmap = SKBitmap.Decode(imageBytes))
            {
                if (bitmap == null || bitmap.Width <= 0 || bitmap.Height <= 0)
                {
                    return false;
                }

                long sumR = 0;
                long sumG = 0;
                long sumB = 0;
                long count = 0;

                for (var y = 0; y < bitmap.Height; y++)
                {
                    for (var x = 0; x < bitmap.Width; x++)
                    {
                        var pixel = bitmap.GetPixel(x, y);
                        if (pixel.Alpha == 0)
                        {
                            continue;
                        }

                        if (pixel.Red > 247 && pixel.Green > 247 && pixel.Blue > 247)
                        {
                            continue;
                        }

                        sumR += pixel.Red;
                        sumG += pixel.Green;
                        sumB += pixel.Blue;
                        count++;
                    }
                }

                if (count == 0)
                {
                    for (var y = 0; y < bitmap.Height; y++)
                    {
                        for (var x = 0; x < bitmap.Width; x++)
                        {
                            var pixel = bitmap.GetPixel(x, y);
                            if (pixel.Alpha == 0)
                            {
                                continue;
                            }

                            sumR += pixel.Red;
                            sumG += pixel.Green;
                            sumB += pixel.Blue;
                            count++;
                        }
                    }
                }

                if (count == 0)
                {
                    return false;
                }

                color = new SKColor(
                    ToByte(sumR / (double)count),
                    ToByte(sumG / (double)count),
                    ToByte(sumB / (double)count));
                return true;
            }
        }

        private static byte ToByte(double value)
        {
            var rounded = Math.Round(value);
            if (rounded < 0d)
            {
                return 0;
            }

            if (rounded > 255d)
            {
                return 255;
            }

            return byte.Parse(rounded.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        }

        private static bool TryResolveSize(XElement defaultRunProperties, out double sizePt)
        {
            sizePt = 0d;
            var raw = AttrOf(defaultRunProperties, "sz");
            double sizeHundredthPoints;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out sizeHundredthPoints))
            {
                return false;
            }

            if (sizeHundredthPoints <= 0d)
            {
                return false;
            }

            sizePt = sizeHundredthPoints / 100d;
            return true;
        }

        private static bool TryResolveBooleanAttr(XElement element, string name, out bool value)
        {
            value = false;
            var raw = AttrOf(element, name);
            if (string.IsNullOrEmpty(raw))
            {
                return false;
            }

            value = !string.Equals(raw, "0", StringComparison.Ordinal)
                && !string.Equals(raw, "false", StringComparison.OrdinalIgnoreCase);
            return true;
        }

        private static bool TryResolveTypeface(XElement defaultRunProperties, out string typeface)
        {
            typeface = null;
            var latin = defaultRunProperties.Element(A + "latin");
            typeface = ResolveTypefaceAttr(latin);
            if (!string.IsNullOrEmpty(typeface))
            {
                return true;
            }

            var eastAsian = defaultRunProperties.Element(A + "ea");
            typeface = ResolveTypefaceAttr(eastAsian);
            if (!string.IsNullOrEmpty(typeface))
            {
                return true;
            }

            var complexScript = defaultRunProperties.Element(A + "cs");
            typeface = ResolveTypefaceAttr(complexScript);
            return !string.IsNullOrEmpty(typeface);
        }

        private static string ResolveTypefaceAttr(XElement element)
        {
            var typeface = AttrOf(element, "typeface");
            if (string.IsNullOrEmpty(typeface))
            {
                return null;
            }

            if (typeface[0] == '+')
            {
                return null;
            }

            return typeface;
        }

        private static string AttrOf(XElement element, XName name)
        {
            if (element == null)
            {
                return null;
            }

            var attribute = element.Attribute(name);
            if (attribute == null)
            {
                return null;
            }

            return attribute.Value;
        }
    }
}
