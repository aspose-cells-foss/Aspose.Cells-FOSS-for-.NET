using System;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class MapChartXmlParser
    {
        private static readonly XNamespace Cx = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static MapChartData Parse(string xml, WorkbookModel workbook)
        {
            if (string.IsNullOrEmpty(xml) || workbook == null)
            {
                return null;
            }

            XDocument document;
            try
            {
                document = XDocument.Parse(xml);
            }
            catch (Exception)
            {
                return null;
            }

            var chart = document.Root != null ? document.Root.Element(Cx + "chart") : null;
            var region = chart != null ? chart.Element(Cx + "plotArea") : null;
            region = region != null ? region.Element(Cx + "plotAreaRegion") : null;
            if (region == null)
            {
                return null;
            }

            XElement visibleSeries = null;
            foreach (var series in region.Elements(Cx + "series"))
            {
                if (string.Equals((string)series.Attribute("layoutId"), "regionMap", StringComparison.Ordinal)
                    && !string.Equals((string)series.Attribute("hidden"), "1", StringComparison.Ordinal))
                {
                    visibleSeries = series;
                    break;
                }
            }

            if (visibleSeries == null)
            {
                return null;
            }

            var dataId = visibleSeries.Element(Cx + "dataId");
            var chartData = document.Root.Element(Cx + "chartData");
            XElement data = null;
            if (dataId != null && chartData != null)
            {
                foreach (var candidate in chartData.Elements(Cx + "data"))
                {
                    if (string.Equals((string)candidate.Attribute("id"), (string)dataId.Attribute("val"), StringComparison.Ordinal))
                    {
                        data = candidate;
                        break;
                    }
                }
            }

            if (data == null)
            {
                return null;
            }

            var categoryDimension = data.Element(Cx + "strDim");
            var numericDimension = data.Element(Cx + "numDim");
            var categoryName = categoryDimension != null ? (string)categoryDimension.Element(Cx + "f") : null;
            var numericName = numericDimension != null ? (string)numericDimension.Element(Cx + "f") : null;
            var categories = ChartWorkbookDataResolver.ResolveStringRange(workbook, ResolveName(workbook, categoryName));
            var values = ChartWorkbookDataResolver.ResolveNumericRange(workbook, ResolveName(workbook, numericName), workbook.Settings.DateSystem);
            var displayValues = ChartWorkbookDataResolver.ResolveStringRange(workbook, ResolveName(workbook, numericName));
            if (categories.Count == 0 || categories.Count != values.Count)
            {
                return null;
            }

            var result = new MapChartData();
            var title = chart.Element(Cx + "title");
            var titleData = title != null ? title.Element(Cx + "tx") : null;
            titleData = titleData != null ? titleData.Element(Cx + "txData") : null;
            result.Title = title != null
                ? (titleData != null ? (string)titleData.Element(Cx + "v") : null) ?? "Chart Title"
                : string.Empty;
            var seriesTitle = visibleSeries.Element(Cx + "tx");
            seriesTitle = seriesTitle != null ? seriesTitle.Element(Cx + "txData") : null;
            result.SeriesName = seriesTitle != null ? (string)seriesTitle.Element(Cx + "v") ?? string.Empty : string.Empty;
            ParseColors(visibleSeries.Element(Cx + "valueColors"), RenderColor.FromWorkbook(workbook), result);

            var hasValues = false;
            for (var i = 0; i < categories.Count; i++)
            {
                if (string.IsNullOrEmpty(categories[i]) || !values[i].HasValue)
                {
                    continue;
                }

                var value = values[i].Value;
                result.Values[categories[i]] = value;
                if (!hasValues)
                {
                    result.Minimum = value;
                    result.Maximum = value;
                    result.MinimumLabel = i < displayValues.Count ? displayValues[i] : string.Empty;
                    result.MaximumLabel = result.MinimumLabel;
                    hasValues = true;
                }
                else
                {
                    if (value < result.Minimum)
                    {
                        result.Minimum = value;
                        result.MinimumLabel = i < displayValues.Count ? displayValues[i] : string.Empty;
                    }
                    if (value > result.Maximum)
                    {
                        result.Maximum = value;
                        result.MaximumLabel = i < displayValues.Count ? displayValues[i] : string.Empty;
                    }
                }
            }

            return hasValues ? result : null;
        }

        private static void ParseColors(XElement valueColors, RenderColor colors, MapChartData result)
        {
            if (valueColors == null)
            {
                return;
            }

            SKColor minimum;
            SKColor maximum;
            if (!TryColor(valueColors.Element(Cx + "minColor"), colors, out minimum)
                || !TryColor(valueColors.Element(Cx + "maxColor"), colors, out maximum))
            {
                return;
            }

            result.HasCustomColors = true;
            result.MinimumColor = minimum;
            result.MaximumColor = maximum;
            SKColor middle;
            if (TryColor(valueColors.Element(Cx + "midColor"), colors, out middle))
            {
                result.HasMiddleColor = true;
                result.MiddleColor = middle;
            }
        }

        private static bool TryColor(XElement container, RenderColor colors, out SKColor color)
        {
            color = SKColors.Transparent;
            if (container == null)
            {
                return false;
            }

            var scheme = container.Element(A + "schemeClr");
            if (scheme != null)
            {
                color = colors.ResolveSchemeName((string)scheme.Attribute("val"), SKColors.Transparent);
                return color.Alpha > 0;
            }

            var rgb = container.Element(A + "srgbClr");
            uint value;
            if (rgb != null && uint.TryParse((string)rgb.Attribute("val"), NumberStyles.HexNumber,
                CultureInfo.InvariantCulture, out value))
            {
                color = new SKColor((byte)(value >> 16), (byte)(value >> 8), (byte)value);
                return true;
            }

            return false;
        }

        private static string ResolveName(WorkbookModel workbook, string formula)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return string.Empty;
            }

            for (var i = 0; i < workbook.DefinedNames.Count; i++)
            {
                var name = workbook.DefinedNames[i];
                if (string.Equals(name.Name, formula, StringComparison.Ordinal))
                {
                    return name.Formula;
                }
            }

            return formula;
        }
    }
}
