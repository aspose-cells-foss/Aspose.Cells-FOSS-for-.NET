using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal enum ChartKind
    {
        Unsupported,
        Line,
        Column,
        Bar,
        Area,
        Radar,
        Pie,
    }

    internal sealed class ChartSeries
    {
        public string Name = string.Empty;
        public SKColor Color = SKColors.Gray;
        public SKColor LineColor = SKColors.Gray;
        public double LineWidthPt = 2.25d;
        public bool HasVisibleLine = true;
        public List<double?> Values = new List<double?>();
        public bool MarkerVisible;
        public string MarkerSymbol = "none";
        public int MarkerSize = 5;
        public SKColor MarkerFillColor = SKColors.Gray;
        public SKColor MarkerStrokeColor = SKColors.Gray;
        public double MarkerStrokeWidthPt = 0.75d;
    }

    /// <summary>
    /// A chart's definition reduced to what the renderer needs, parsed from the cached values in the
    /// chart XML (no formula engine required).
    /// </summary>
    internal sealed class ParsedChart
    {
        public ChartKind Kind = ChartKind.Unsupported;
        public CultureInfo Culture;
        public string Title;                 // null = no title band
        public FontValue TitleFont = new FontValue { Name = "Calibri", Size = 14d, Bold = true };
        public SKColor TitleColor = new SKColor(0x59, 0x59, 0x59);
        public bool HasTitleBorder;
        public SKColor TitleBorderColor = SKColors.Transparent;
        public double TitleBorderWidthPt = 0.75d;
        public SKColor AxisLabelColor = new SKColor(0x59, 0x59, 0x59);
        public SKColor LegendTextColor = new SKColor(0x59, 0x59, 0x59);
        public SKColor SeriesAxisLabelColor = new SKColor(0x59, 0x59, 0x59);
        public bool HasChartAreaFill;
        public SKColor ChartAreaFillColor = SKColors.White;
        public string LegendPosition;        // "b","t","l","r","tr", or null
        public bool HasValueGridlines;
        public bool HasMinorValueGridlines;
        public string ValueFormatCode;
        public bool CrossBetween = true;  // category points sit at band centers (Excel default)
        public bool IsStacked;
        public bool IsPercentStacked;
        public double? AxisMin;              // explicit scaling, else auto
        public double? AxisMax;
        public double? MajorUnit;
        public double? MinorUnit;
        public SKColor MinorValueGridlineColor = new SKColor(0xEDEDED);
        public double MinorValueGridlineWidthPt = 1d;
        public bool HasCategoryGridlines;
        public SKColor CategoryGridlineColor = new SKColor(0xD9, 0xD9, 0xD9);
        public double CategoryGridlineWidthPt = 1d;
        public bool HasPlotAreaBorder;
        public SKColor PlotAreaBorderColor = SKColors.Transparent;
        public double PlotAreaBorderWidthPt = 1d;
        public List<ChartSeries> Series = new List<ChartSeries>();
        public List<string> Categories = new List<string>();
        public List<double?> CategoryValues = new List<double?>();
        public List<SKColor> PiePointColors = new List<SKColor>();
        public bool IsOfPie;
        public string OfPieType;
        public string OfPieSplitType;
        public double? OfPieSplitPosition;
        public double OfPieSecondPieSizePercent = 75d;
        public double OfPieGapWidthPercent = 150d;
        public bool HasOfPieSeparatorLines;
        public SKColor OfPieSeparatorLineColor = new SKColor(0xA6, 0xA6, 0xA6);
        public double OfPieSeparatorLineWidthPt = 0.75d;
        public bool HasDateCategoryAxis;
        public string CategoryFormatCode;
        public string CategoryBaseTimeUnit;
        public double CategoryAxisTextRotationDeg;
        public bool CategoryAxisMinMax = true;
        public string RadarStyle = "standard";
        public double GapWidthPercent = 150d;
        public double OverlapPercent;
        public bool IsLine3D;
        public double RotationXDeg;
        public double RotationYDeg;
        public double PerspectiveDeg;
        public bool HasVisibleSeriesAxis;

        // Manual inner-plot layout as fractions of the chart area (Excel c:manualLayout), if present.
        public bool HasManualPlot;
        public double PlotX, PlotY, PlotW, PlotH;
    }

    internal static class ChartXmlParser
    {
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static ParsedChart Parse(string xml, RenderColor colors, DateSystem dateSystem, CultureInfo culture, WorkbookModel workbook = null, ChartModel chartModel = null)
        {
            if (string.IsNullOrEmpty(xml))
            {
                return null;
            }

            XDocument doc;
            try { doc = XDocument.Parse(xml); }
            catch (Exception) { return null; }

            var chart = doc.Root != null ? doc.Root.Element(C + "chart") : null;
            var plotArea = chart != null ? chart.Element(C + "plotArea") : null;
            if (plotArea == null)
            {
                return null;
            }

            ChartKind kind;
            XElement typeEl = FindTypeElement(plotArea, out kind);
            if (typeEl == null)
            {
                return null;
            }

            var result = new ParsedChart { Kind = kind };
            result.Culture = ResolveChartCulture(doc.Root, culture);
            result.Title = ParseTitle(chart, colors, result, chartModel, workbook);
            ParseChartAreaStyle(doc.Root, colors, result);
            ParseLegend(chart, colors, result, chartModel);
            ParseManualPlot(plotArea, result);
            ParseThreeDimensionalSettings(chart, typeEl, result);
            ParseOfPieSettings(typeEl, colors, result);
            ParseGrouping(typeEl, result);
            ParseRadarSettings(typeEl, result);
            ParseBarSeriesLayout(typeEl, result);
            ParseValueAxis(plotArea, colors, result, chartModel);
            ParseCategoryAxis(plotArea, colors, result, chartModel);
            ParseSeriesAxis(plotArea, colors, result, chartModel);
            ParsePlotAreaStyle(plotArea, colors, result);

            var seriesIndex = 0;
            foreach (var ser in typeEl.Elements(C + "ser"))
            {
                var series = new ChartSeries
                {
                    Name = ParseSeriesName(ser.Element(C + "tx"), workbook),
                    Color = ParseSeriesColor(ser, colors, seriesIndex),
                    LineWidthPt = ParseLineWidthPt(ser),
                    Values = ParseNumericData(ser.Element(C + "val"), workbook, dateSystem),
                };
                ParseSeriesLineStyle(ser, colors, series);
                ParseMarker(ser, colors, series);
                ApplyRadarSeriesDefaults(result, series);
                result.Series.Add(series);
                seriesIndex++;
            }

            ApplyFallbackTitle(result);

            var firstSer = FirstElement(typeEl, C + "ser");
            if (firstSer != null)
            {
                result.Categories = ParseCategories(firstSer.Element(C + "cat"), workbook, dateSystem, result.Culture);
                result.CategoryValues = ParseCategoryValues(firstSer.Element(C + "cat"), workbook, dateSystem);
                result.PiePointColors = ParsePiePointColors(firstSer, colors, result.Categories.Count);
                if (string.IsNullOrEmpty(result.ValueFormatCode))
                {
                    result.ValueFormatCode = TextOf(FirstDescendant(firstSer.Element(C + "val"), C + "formatCode"));
                }
            }

            return result;
        }

        private static void ApplyFallbackTitle(ParsedChart chart)
        {
            if (chart == null)
            {
                return;
            }

            if (chart.Kind == ChartKind.Pie
                && chart.Series != null
                && chart.Series.Count > 0
                && (string.IsNullOrEmpty(chart.Title) || string.Equals(chart.Title, "Chart Title", StringComparison.Ordinal)))
            {
                chart.Title = chart.Series[0].Name;
            }
        }

        private static CultureInfo ResolveChartCulture(XElement chartSpace, CultureInfo fallbackCulture)
        {
            var cultureName = AttrOf(chartSpace != null ? chartSpace.Element(C + "lang") : null, "val");
            if (string.IsNullOrEmpty(cultureName))
            {
                return fallbackCulture ?? CultureInfo.CurrentCulture;
            }

            try
            {
                return CultureInfo.GetCultureInfo(cultureName);
            }
            catch (CultureNotFoundException)
            {
                return fallbackCulture ?? CultureInfo.CurrentCulture;
            }
        }

        private static XElement FindTypeElement(XElement plotArea, out ChartKind kind)
        {
            var line = plotArea.Element(C + "lineChart");
            if (line != null) { kind = ChartKind.Line; return line; }

            var line3d = plotArea.Element(C + "line3DChart");
            if (line3d != null) { kind = ChartKind.Line; return line3d; }

            var bar = plotArea.Element(C + "barChart");
            if (bar != null)
            {
                var dir = AttrOf(bar.Element(C + "barDir"), "val");
                kind = string.Equals(dir, "bar", StringComparison.Ordinal) ? ChartKind.Bar : ChartKind.Column;
                return bar;
            }

            var bar3d = plotArea.Element(C + "bar3DChart");
            if (bar3d != null)
            {
                var dir = AttrOf(bar3d.Element(C + "barDir"), "val");
                kind = string.Equals(dir, "bar", StringComparison.Ordinal) ? ChartKind.Bar : ChartKind.Column;
                return bar3d;
            }

            var area = plotArea.Element(C + "areaChart");
            if (area != null) { kind = ChartKind.Area; return area; }

            var area3d = plotArea.Element(C + "area3DChart");
            if (area3d != null) { kind = ChartKind.Area; return area3d; }

            var radar = plotArea.Element(C + "radarChart");
            if (radar != null) { kind = ChartKind.Radar; return radar; }

            var pie = plotArea.Element(C + "pieChart");
            if (pie != null) { kind = ChartKind.Pie; return pie; }

            var pie3d = plotArea.Element(C + "pie3DChart");
            if (pie3d != null) { kind = ChartKind.Pie; return pie3d; }

            var ofPie = plotArea.Element(C + "ofPieChart");
            if (ofPie != null) { kind = ChartKind.Pie; return ofPie; }

            kind = ChartKind.Unsupported;
            return null;
        }

        private static string ParseTitle(XElement chart, RenderColor colors, ParsedChart result, ChartModel chartModel, WorkbookModel workbook)
        {
            var titleEl = chart.Element(C + "title");
            if (titleEl == null)
            {
                return null;
            }

            if (string.Equals(AttrOf(chart.Element(C + "autoTitleDeleted"), "val"), "1", StringComparison.Ordinal))
            {
                return null;
            }

            ParseTitleStyle(titleEl, colors, result, chartModel);
            var text = ParseTitleText(titleEl, workbook);
            if (!string.IsNullOrEmpty(text))
            {
                return text;
            }

            return "Chart Title";
        }

        private static void ParseTitleStyle(XElement titleEl, RenderColor colors, ParsedChart result, ChartModel chartModel)
        {
            if (titleEl == null || result == null)
            {
                return;
            }

            var titleTextProperties = ResolveTextProperties(titleEl);

            SKColor titleColor;
            if (TryResolveTextColor(titleTextProperties, colors, chartModel, out titleColor))
            {
                result.TitleColor = titleColor;
            }

            FontValue titleFont;
            if (TryResolveTextFont(titleTextProperties, out titleFont))
            {
                result.TitleFont = titleFont;
            }

            var shapeProperties = titleEl.Element(C + "spPr");
            var line = shapeProperties != null ? shapeProperties.Element(A + "ln") : null;
            ParseLineStyle(line, colors, out result.HasTitleBorder, out result.TitleBorderColor, out result.TitleBorderWidthPt);
            if (result.HasTitleBorder && line != null && line.Attribute("w") == null)
            {
                result.TitleBorderWidthPt = 0.75d;
            }
        }

        private static void ParseLegend(XElement chart, RenderColor colors, ParsedChart result, ChartModel chartModel)
        {
            var legend = chart.Element(C + "legend");
            if (legend != null)
            {
                result.LegendPosition = AttrOf(legend.Element(C + "legendPos"), "val") ?? "r";

                SKColor legendTextColor;
                if (TryResolveTextColor(legend.Element(C + "txPr"), colors, chartModel, out legendTextColor))
                {
                    result.LegendTextColor = legendTextColor;
                }
            }
        }

        private static void ParseGrouping(XElement typeEl, ParsedChart result)
        {
            if (typeEl == null || result == null)
            {
                return;
            }

            var grouping = AttrOf(typeEl.Element(C + "grouping"), "val");
            result.IsStacked = string.Equals(grouping, "stacked", StringComparison.Ordinal);
            result.IsPercentStacked = string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        }

        private static void ParseBarSeriesLayout(XElement typeEl, ParsedChart result)
        {
            if (typeEl == null || result == null)
            {
                return;
            }

            if (typeEl.Name != C + "barChart" && typeEl.Name != C + "bar3DChart")
            {
                return;
            }

            double gapWidth;
            if (TryDouble(AttrOf(typeEl.Element(C + "gapWidth"), "val"), out gapWidth) && gapWidth >= 0d)
            {
                result.GapWidthPercent = gapWidth;
            }

            double overlap;
            if (TryDouble(AttrOf(typeEl.Element(C + "overlap"), "val"), out overlap))
            {
                if (overlap < -100d)
                {
                    overlap = -100d;
                }
                else if (overlap > 100d)
                {
                    overlap = 100d;
                }

                result.OverlapPercent = overlap;
            }
        }

        private static void ParseRadarSettings(XElement typeEl, ParsedChart result)
        {
            if (typeEl == null || result == null)
            {
                return;
            }

            if (typeEl.Name != C + "radarChart")
            {
                return;
            }

            var radarStyle = AttrOf(typeEl.Element(C + "radarStyle"), "val");
            if (!string.IsNullOrEmpty(radarStyle))
            {
                result.RadarStyle = radarStyle;
            }
        }

        private static void ParseOfPieSettings(XElement typeEl, RenderColor colors, ParsedChart result)
        {
            if (typeEl == null || result == null)
            {
                return;
            }

            if (typeEl.Name != C + "ofPieChart")
            {
                return;
            }

            result.IsOfPie = true;
            result.OfPieType = AttrOf(typeEl.Element(C + "ofPieType"), "val") ?? "pie";
            result.OfPieSplitType = AttrOf(typeEl.Element(C + "splitType"), "val");

            double value;
            if (TryDouble(AttrOf(typeEl.Element(C + "splitPos"), "val"), out value))
            {
                result.OfPieSplitPosition = value;
            }

            if (TryDouble(AttrOf(typeEl.Element(C + "secondPieSize"), "val"), out value) && value > 0d)
            {
                result.OfPieSecondPieSizePercent = value;
            }

            if (TryDouble(AttrOf(typeEl.Element(C + "gapWidth"), "val"), out value) && value >= 0d)
            {
                result.OfPieGapWidthPercent = value;
            }

            var separatorLine = typeEl.Element(C + "serLines");
            if (separatorLine == null)
            {
                return;
            }

            var line = separatorLine.Element(C + "spPr") != null ? separatorLine.Element(C + "spPr").Element(A + "ln") : null;
            SKColor color;
            double widthPt;
            bool hasVisibleLine;
            ParseLineStyle(line, colors, out hasVisibleLine, out color, out widthPt);
            result.HasOfPieSeparatorLines = hasVisibleLine || line == null;
            if (hasVisibleLine)
            {
                result.OfPieSeparatorLineColor = color;
                result.OfPieSeparatorLineWidthPt = widthPt;
            }
        }

        private static void ParseThreeDimensionalSettings(XElement chart, XElement typeEl, ParsedChart result)
        {
            if (chart == null || typeEl == null || result == null)
            {
                return;
            }

            if (typeEl.Name != C + "line3DChart")
            {
                return;
            }

            result.IsLine3D = true;
            var view3D = chart.Element(C + "view3D");
            double value;
            if (TryDouble(AttrOf(view3D != null ? view3D.Element(C + "rotX") : null, "val"), out value))
            {
                result.RotationXDeg = value;
            }

            if (TryDouble(AttrOf(view3D != null ? view3D.Element(C + "rotY") : null, "val"), out value))
            {
                result.RotationYDeg = value;
            }

            if (TryDouble(AttrOf(view3D != null ? view3D.Element(C + "perspective") : null, "val"), out value))
            {
                result.PerspectiveDeg = value;
            }
        }

        private static void ParseManualPlot(XElement plotArea, ParsedChart result)
        {
            var manual = plotArea.Element(C + "layout") != null
                ? plotArea.Element(C + "layout").Element(C + "manualLayout")
                : null;
            if (manual == null || !string.Equals(AttrOf(manual.Element(C + "layoutTarget"), "val"), "inner", StringComparison.Ordinal))
            {
                return;
            }

            double x, y, w, h;
            if (TryDouble(AttrOf(manual.Element(C + "x"), "val"), out x)
                && TryDouble(AttrOf(manual.Element(C + "y"), "val"), out y)
                && TryDouble(AttrOf(manual.Element(C + "w"), "val"), out w)
                && TryDouble(AttrOf(manual.Element(C + "h"), "val"), out h))
            {
                result.HasManualPlot = true;
                result.PlotX = x; result.PlotY = y; result.PlotW = w; result.PlotH = h;
            }
        }

        private static void ParseValueAxis(XElement plotArea, RenderColor colors, ParsedChart result, ChartModel chartModel)
        {
            var valAx = FirstElement(plotArea, C + "valAx");
            if (valAx == null)
            {
                return;
            }

            result.HasValueGridlines = valAx.Element(C + "majorGridlines") != null;
            result.HasMinorValueGridlines = valAx.Element(C + "minorGridlines") != null;
            result.ValueFormatCode = AttrOf(valAx.Element(C + "numFmt"), "formatCode");
            result.CrossBetween = !string.Equals(AttrOf(valAx.Element(C + "crossBetween"), "val"), "midCat", StringComparison.Ordinal);

            ParseValueGridlineStyle(valAx.Element(C + "minorGridlines"), colors, result);
            var scaling = valAx.Element(C + "scaling");
            double v;
            if (scaling != null)
            {
                if (TryDouble(AttrOf(scaling.Element(C + "min"), "val"), out v)) result.AxisMin = v;
                if (TryDouble(AttrOf(scaling.Element(C + "max"), "val"), out v)) result.AxisMax = v;
            }
            if (TryDouble(AttrOf(valAx.Element(C + "majorUnit"), "val"), out v)) result.MajorUnit = v;
            if (TryDouble(AttrOf(valAx.Element(C + "minorUnit"), "val"), out v)) result.MinorUnit = v;

            SKColor axisLabelColor;
            if (TryResolveTextColor(valAx.Element(C + "txPr"), colors, chartModel, out axisLabelColor))
            {
                result.AxisLabelColor = axisLabelColor;
            }
        }

        private static void ParseCategoryAxis(XElement plotArea, RenderColor colors, ParsedChart result, ChartModel chartModel)
        {
            var dateAxis = FirstElement(plotArea, C + "dateAx");
            var categoryAxis = FirstElement(plotArea, C + "catAx");
            var axisElement = dateAxis ?? categoryAxis;
            if (axisElement == null)
            {
                return;
            }

            result.HasDateCategoryAxis = dateAxis != null;
            result.CategoryFormatCode = AttrOf(axisElement.Element(C + "numFmt"), "formatCode");
            result.CategoryBaseTimeUnit = AttrOf(axisElement.Element(C + "baseTimeUnit"), "val");
            var scaling = axisElement.Element(C + "scaling");
            result.CategoryAxisMinMax = !string.Equals(AttrOf(scaling != null ? scaling.Element(C + "orientation") : null, "val"), "maxMin", StringComparison.Ordinal);
            ParseGridlineStyle(axisElement.Element(C + "majorGridlines"), colors, result);

            var bodyPr = axisElement.Element(C + "txPr") != null ? axisElement.Element(C + "txPr").Element(A + "bodyPr") : null;
            double rotation;
            if (TryDouble(AttrOf(bodyPr, "rot"), out rotation))
            {
                result.CategoryAxisTextRotationDeg = rotation / 1000000d;
            }

            SKColor axisLabelColor;
            if (TryResolveTextColor(axisElement.Element(C + "txPr"), colors, chartModel, out axisLabelColor))
            {
                result.AxisLabelColor = axisLabelColor;
            }
        }

        private static void ParsePlotAreaStyle(XElement plotArea, RenderColor colors, ParsedChart result)
        {
            if (plotArea == null || colors == null || result == null)
            {
                return;
            }

            var shapeProperties = plotArea.Element(C + "spPr");
            var line = shapeProperties != null ? shapeProperties.Element(A + "ln") : null;
            ParseLineStyle(line, colors, out result.HasPlotAreaBorder, out result.PlotAreaBorderColor, out result.PlotAreaBorderWidthPt);
        }

        private static void ParseChartAreaStyle(XElement chartSpace, RenderColor colors, ParsedChart result)
        {
            if (chartSpace == null || colors == null || result == null)
            {
                return;
            }

            var shapeProperties = chartSpace.Element(C + "spPr");
            var solidFill = shapeProperties != null ? shapeProperties.Element(A + "solidFill") : null;
            var fill = FillChild(solidFill);
            if (fill == null)
            {
                return;
            }

            SKColor color;
            if (TryDrawingColor(fill, colors, out color))
            {
                result.HasChartAreaFill = true;
                result.ChartAreaFillColor = color;
            }
        }

        private static void ParseSeriesAxis(XElement plotArea, RenderColor colors, ParsedChart result, ChartModel chartModel)
        {
            var serAx = FirstElement(plotArea, C + "serAx");
            if (serAx == null)
            {
                return;
            }

            result.HasVisibleSeriesAxis = !string.Equals(AttrOf(serAx.Element(C + "delete"), "val"), "1", StringComparison.Ordinal);
            SKColor axisLabelColor;
            if (TryResolveTextColor(serAx.Element(C + "txPr"), colors, chartModel, out axisLabelColor))
            {
                result.SeriesAxisLabelColor = axisLabelColor;
            }
        }

        private static void ParseGridlineStyle(XElement gridlines, RenderColor colors, ParsedChart result)
        {
            if (gridlines == null || result == null)
            {
                return;
            }

            var line = gridlines.Element(C + "spPr") != null ? gridlines.Element(C + "spPr").Element(A + "ln") : null;
            SKColor color;
            double widthPt;
            bool hasVisibleLine;
            ParseLineStyle(line, colors, out hasVisibleLine, out color, out widthPt);
            result.HasCategoryGridlines = hasVisibleLine || line == null;
            if (hasVisibleLine)
            {
                result.CategoryGridlineColor = color;
                result.CategoryGridlineWidthPt = widthPt;
            }
        }

        private static void ParseValueGridlineStyle(XElement gridlines, RenderColor colors, ParsedChart result)
        {
            if (gridlines == null || result == null)
            {
                return;
            }

            var line = gridlines.Element(C + "spPr") != null ? gridlines.Element(C + "spPr").Element(A + "ln") : null;
            SKColor color;
            double widthPt;
            bool hasVisibleLine;
            ParseLineStyle(line, colors, out hasVisibleLine, out color, out widthPt);
            result.HasMinorValueGridlines = hasVisibleLine || line == null;
            if (hasVisibleLine)
            {
                result.MinorValueGridlineColor = color;
                result.MinorValueGridlineWidthPt = widthPt;
            }
        }

        private static void ParseLineStyle(XElement line, RenderColor colors, out bool hasVisibleLine, out SKColor color, out double widthPt)
        {
            hasVisibleLine = false;
            color = new SKColor(0xD9, 0xD9, 0xD9);
            widthPt = 1d;

            if (line == null)
            {
                return;
            }

            if (line.Element(A + "noFill") != null)
            {
                return;
            }

            double emu;
            if (TryDouble(AttrOf(line, "w"), out emu) && emu > 0d)
            {
                widthPt = emu / 12700d;
            }

            var solidFill = line.Element(A + "solidFill");
            var fill = FillChild(solidFill);
            if (fill != null)
            {
                SKColor parsedColor;
                if (TryDrawingColor(fill, colors, out parsedColor))
                {
                    color = parsedColor;
                    hasVisibleLine = true;
                    return;
                }
            }

            hasVisibleLine = line.Element(A + "solidFill") == null;
        }

        internal static bool TryDrawingColor(XElement source, RenderColor colors, out SKColor color)
        {
            color = SKColors.Transparent;
            if (source == null || colors == null)
            {
                return false;
            }

            if (source.Name == A + "srgbClr")
            {
                return TryHexWithAlpha(AttrOf(source, "val"), source, out color);
            }

            if (source.Name == A + "schemeClr")
            {
                var baseColor = colors.ResolveSchemeName(AttrOf(source, "val"), SKColors.Transparent);
                double lumMod;
                double lumOff;
                TryDouble(AttrOf(source.Element(A + "lumMod"), "val"), out lumMod);
                TryDouble(AttrOf(source.Element(A + "lumOff"), "val"), out lumOff);
                if (lumMod > 0d || lumOff > 0d)
                {
                    baseColor = RenderColor.ApplyLuma(baseColor, lumMod > 0d ? lumMod / 100000d : 1d, lumOff / 100000d);
                }

                double alpha;
                if (TryDouble(AttrOf(source.Element(A + "alpha"), "val"), out alpha) && alpha >= 0d)
                {
                    var channel = (byte)Math.Round(Math.Min(100000d, alpha) / 100000d * 255d);
                    color = new SKColor(baseColor.Red, baseColor.Green, baseColor.Blue, channel);
                }
                else
                {
                    color = baseColor;
                }

                return true;
            }

            return false;
        }

        private static bool TryHexWithAlpha(string hex, XElement source, out SKColor color)
        {
            color = SKColors.Transparent;
            SKColor rgb;
            if (!TryHex(hex, out rgb))
            {
                return false;
            }

            double alpha;
            if (TryDouble(AttrOf(source.Element(A + "alpha"), "val"), out alpha) && alpha >= 0d)
            {
                var channel = (byte)Math.Round(Math.Min(100000d, alpha) / 100000d * 255d);
                color = new SKColor(rgb.Red, rgb.Green, rgb.Blue, channel);
            }
            else
            {
                color = rgb;
            }

            return true;
        }

        private static SKColor ParseSeriesColor(XElement ser, RenderColor colors, int index)
        {
            var line = ser.Element(C + "spPr") != null ? ser.Element(C + "spPr").Element(A + "ln") : null;
            var fill = ser.Element(C + "spPr") != null ? ser.Element(C + "spPr").Element(A + "solidFill") : null;
            var source = FirstNonNull(FillChild(fill), FillChild(line != null ? line.Element(A + "solidFill") : null));

            if (source != null)
            {
                if (source.Name == A + "srgbClr")
                {
                    SKColor parsed;
                    if (TryHex(AttrOf(source, "val"), out parsed)) return parsed;
                }
                else if (source.Name == A + "schemeClr")
                {
                    var name = AttrOf(source, "val");
                    var baseColor = colors.ResolveSchemeName(name, DefaultSeriesColor(colors, index));
                    double lumMod, lumOff;
                    TryDouble(AttrOf(source.Element(A + "lumMod"), "val"), out lumMod);
                    TryDouble(AttrOf(source.Element(A + "lumOff"), "val"), out lumOff);
                    if (lumMod > 0d || lumOff > 0d)
                    {
                        return RenderColor.ApplyLuma(baseColor, lumMod > 0d ? lumMod / 100000d : 1d, lumOff / 100000d);
                    }
                    return baseColor;
                }
            }

            return DefaultSeriesColor(colors, index);
        }

        internal static XElement FillChild(XElement solidFill)
        {
            if (solidFill == null) return null;
            return solidFill.Element(A + "srgbClr") ?? solidFill.Element(A + "schemeClr");
        }

        private static List<SKColor> ParsePiePointColors(XElement ser, RenderColor colors, int pointCount)
        {
            var result = new List<SKColor>();
            if (pointCount < 0)
            {
                pointCount = 0;
            }

            for (var i = 0; i < pointCount + 1; i++)
            {
                result.Add(DefaultSeriesColor(colors, i));
            }

            if (ser == null)
            {
                return result;
            }

            foreach (var point in ser.Elements(C + "dPt"))
            {
                int index;
                if (!int.TryParse(AttrOf(point.Element(C + "idx"), "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out index))
                {
                    continue;
                }

                while (index >= result.Count)
                {
                    result.Add(DefaultSeriesColor(colors, result.Count));
                }

                var shapeProperties = point.Element(C + "spPr");
                var fill = shapeProperties != null ? shapeProperties.Element(A + "solidFill") : null;
                var source = FillChild(fill);
                SKColor color;
                if (TryDrawingColor(source, colors, out color))
                {
                    result[index] = color;
                }
            }

            return result;
        }

        private static bool TryResolveTextColor(XElement textProperties, RenderColor colors, ChartModel chartModel, out SKColor color)
        {
            color = SKColors.Transparent;
            var defaultRunProperties = ResolveDefaultRunProperties(textProperties);
            if (defaultRunProperties == null)
            {
                return false;
            }

            if (ChartTextStyleResolver.TryResolveTextColor(defaultRunProperties, colors, chartModel, out color))
            {
                return true;
            }

            var paragraph = textProperties.Element(A + "p");
            var endParagraphRunProperties = paragraph != null ? paragraph.Element(A + "endParaRPr") : null;
            return ChartTextStyleResolver.TryResolveTextColor(endParagraphRunProperties, colors, chartModel, out color);
        }

        private static bool TryResolveTextFont(XElement textProperties, out FontValue font)
        {
            font = null;
            var defaultRunProperties = ResolveDefaultRunProperties(textProperties);
            if (defaultRunProperties == null)
            {
                return false;
            }

            return ChartTextStyleResolver.TryResolveFont(defaultRunProperties, out font);
        }

        private static XElement ResolveTextProperties(XElement owner)
        {
            if (owner == null)
            {
                return null;
            }

            var textProperties = owner.Element(C + "txPr");
            if (textProperties != null)
            {
                return textProperties;
            }

            var richText = owner.Element(C + "tx");
            if (richText != null)
            {
                richText = richText.Element(C + "rich");
            }

            return richText;
        }

        private static XElement ResolveDefaultRunProperties(XElement textProperties)
        {
            if (textProperties == null)
            {
                return null;
            }

            var paragraph = textProperties.Element(A + "p");
            var paragraphProperties = paragraph != null ? paragraph.Element(A + "pPr") : null;
            return paragraphProperties != null ? paragraphProperties.Element(A + "defRPr") : null;
        }

        private static SKColor DefaultSeriesColor(RenderColor colors, int index)
        {
            // Cycle the Office accent slots when a series has no explicit color.
            return colors.ResolveSchemeName("accent" + ((index % 6) + 1), SKColors.Gray);
        }

        private static double ParseLineWidthPt(XElement ser)
        {
            var line = ser.Element(C + "spPr") != null ? ser.Element(C + "spPr").Element(A + "ln") : null;
            double emu;
            if (line != null && TryDouble(AttrOf(line, "w"), out emu) && emu > 0d)
            {
                return emu / 12700d;
            }
            return 2.25d;
        }

        private static void ParseSeriesLineStyle(XElement ser, RenderColor colors, ChartSeries series)
        {
            if (ser == null || colors == null || series == null)
            {
                return;
            }

            series.LineColor = series.Color;
            series.HasVisibleLine = true;

            var line = ser.Element(C + "spPr") != null ? ser.Element(C + "spPr").Element(A + "ln") : null;
            if (line == null)
            {
                return;
            }

            if (line.Element(A + "noFill") != null)
            {
                series.HasVisibleLine = false;
                return;
            }

            SKColor color;
            double widthPt;
            bool hasVisibleLine;
            ParseLineStyle(line, colors, out hasVisibleLine, out color, out widthPt);
            series.HasVisibleLine = hasVisibleLine;
            if (hasVisibleLine)
            {
                series.LineColor = color;
                series.LineWidthPt = widthPt;
            }
        }

        private static void ParseMarker(XElement ser, RenderColor colors, ChartSeries series)
        {
            if (ser == null || series == null)
            {
                return;
            }

            series.MarkerFillColor = series.Color;
            series.MarkerStrokeColor = series.Color;
            series.MarkerStrokeWidthPt = 0.75d;

            var marker = ser.Element(C + "marker");
            if (marker == null)
            {
                return;
            }

            var symbol = AttrOf(marker.Element(C + "symbol"), "val");
            if (!string.IsNullOrEmpty(symbol))
            {
                series.MarkerSymbol = symbol;
                series.MarkerVisible = !string.Equals(symbol, "none", StringComparison.Ordinal);
            }

            int size;
            if (int.TryParse(AttrOf(marker.Element(C + "size"), "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out size) && size > 0)
            {
                series.MarkerSize = size;
            }

            var markerShape = marker.Element(C + "spPr");
            var fill = markerShape != null ? markerShape.Element(A + "solidFill") : null;
            var fillSource = FillChild(fill);
            SKColor markerFillColor;
            if (TryDrawingColor(fillSource, colors, out markerFillColor))
            {
                series.MarkerFillColor = markerFillColor;
            }

            var line = markerShape != null ? markerShape.Element(A + "ln") : null;
            if (line != null)
            {
                SKColor markerStrokeColor;
                double markerStrokeWidth;
                bool hasVisibleLine;
                ParseLineStyle(line, colors, out hasVisibleLine, out markerStrokeColor, out markerStrokeWidth);
                if (hasVisibleLine)
                {
                    series.MarkerStrokeColor = markerStrokeColor;
                    series.MarkerStrokeWidthPt = markerStrokeWidth;
                }
            }
        }

        private static void ApplyRadarSeriesDefaults(ParsedChart chart, ChartSeries series)
        {
            if (chart == null || series == null || chart.Kind != ChartKind.Radar)
            {
                return;
            }

        }

        private static List<double?> ParseNumericData(XElement holder, WorkbookModel workbook, DateSystem dateSystem)
        {
            var formula = TextOf(FirstDescendant(holder, C + "f"));
            var resolved = ChartWorkbookDataResolver.ResolveNumericRange(workbook, formula, dateSystem);
            if (resolved.Count > 0)
            {
                return resolved;
            }

            return ParseNumericCache(holder);
        }

        private static List<double?> ParseNumericCache(XElement holder)
        {
            var result = new List<double?>();
            if (holder == null)
            {
                return result;
            }

            var cache = Descendant(holder, C + "numCache") ?? Descendant(holder, C + "numLit");
            if (cache == null)
            {
                return result;
            }

            int count;
            int.TryParse(AttrOf(cache.Element(C + "ptCount"), "val"), out count);
            var values = new double?[Math.Max(0, count)];
            foreach (var pt in cache.Elements(C + "pt"))
            {
                int idx;
                double v;
                if (int.TryParse(AttrOf(pt, "idx"), out idx) && idx >= 0 && idx < values.Length
                    && TryDouble(TextOf(pt.Element(C + "v")), out v))
                {
                    values[idx] = v;
                }
            }
            result.AddRange(values);
            return result;
        }

        private static List<string> ParseCategories(XElement catEl, WorkbookModel workbook, DateSystem dateSystem, CultureInfo culture)
        {
            var labels = new List<string>();
            if (catEl == null)
            {
                return labels;
            }

            var stringFormula = TextOf(FirstDescendant(catEl.Element(C + "strRef"), C + "f"));
            if (!string.IsNullOrEmpty(stringFormula))
            {
                var resolvedStringLabels = ChartWorkbookDataResolver.ResolveStringRange(workbook, stringFormula);
                if (resolvedStringLabels.Count > 0)
                {
                    return resolvedStringLabels;
                }
            }

            var numericFormula = TextOf(FirstDescendant(catEl.Element(C + "numRef"), C + "f"));
            var numericFormat = TextOf(FirstDescendant(catEl, C + "formatCode"));
            var resolvedLabels = ChartWorkbookDataResolver.ResolveCategoryLabels(workbook, numericFormula, numericFormat, dateSystem, culture);
            if (resolvedLabels.Count > 0)
            {
                return resolvedLabels;
            }

            var strCache = Descendant(catEl, C + "strCache") ?? Descendant(catEl, C + "strLit");
            if (strCache != null)
            {
                int count;
                int.TryParse(AttrOf(strCache.Element(C + "ptCount"), "val"), out count);
                var arr = new string[Math.Max(0, count)];
                foreach (var pt in strCache.Elements(C + "pt"))
                {
                    int idx;
                    if (int.TryParse(AttrOf(pt, "idx"), out idx) && idx >= 0 && idx < arr.Length)
                    {
                        arr[idx] = TextOf(pt.Element(C + "v"));
                    }
                }
                for (var i = 0; i < arr.Length; i++) labels.Add(arr[i] ?? string.Empty);
                return labels;
            }

            var numCache = Descendant(catEl, C + "numCache");
            if (numCache != null)
            {
                var code = TextOf(numCache.Element(C + "formatCode"));
                int count;
                int.TryParse(AttrOf(numCache.Element(C + "ptCount"), "val"), out count);
                var arr = new string[Math.Max(0, count)];
                foreach (var pt in numCache.Elements(C + "pt"))
                {
                    int idx;
                    double v;
                    if (int.TryParse(AttrOf(pt, "idx"), out idx) && idx >= 0 && idx < arr.Length
                        && TryDouble(TextOf(pt.Element(C + "v")), out v))
                    {
                        arr[idx] = FormatLabel(v, code, dateSystem, culture);
                    }
                }
                for (var i = 0; i < arr.Length; i++) labels.Add(arr[i] ?? string.Empty);
            }

            return labels;
        }

        private static List<double?> ParseCategoryValues(XElement catEl, WorkbookModel workbook, DateSystem dateSystem)
        {
            return ParseNumericData(catEl, workbook, dateSystem);
        }

        /// <summary>Formats a numeric axis/category value with its cached format code.</summary>
        public static string FormatLabel(double value, string formatCode, DateSystem dateSystem, CultureInfo culture)
        {
            if (string.IsNullOrEmpty(formatCode) || string.Equals(formatCode, "General", StringComparison.Ordinal))
            {
                return value.ToString("G15", CultureInfo.InvariantCulture);
            }

            var style = new StyleValue { NumberFormat = new NumberFormatValue { Custom = formatCode } };
            object boxed = value;
            if (IsDateFormat(formatCode))
            {
                boxed = DateSerialConverter.FromSerial(value, dateSystem);
            }

            try
            {
                return DisplayTextFormatter.FormatDisplayValue(boxed, style, culture ?? CultureInfo.CurrentCulture);
            }
            catch (Exception)
            {
                return value.ToString("G15", CultureInfo.InvariantCulture);
            }
        }

        private static bool IsDateFormat(string code)
        {
            // Ignore color/condition blocks like [Red] (whose letters would otherwise look like date
            // tokens), backslash escapes, and quoted literals before checking for real date tokens.
            var cleaned = System.Text.RegularExpressions.Regex.Replace(code, @"\[[^\]]*\]", string.Empty);
            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, @"\\.", string.Empty);
            cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "\"[^\"]*\"", string.Empty);

            foreach (var ch in cleaned)
            {
                if (ch == 'y' || ch == 'd' || ch == 'h' || ch == 's')
                {
                    return true;
                }
            }

            return false;
        }

        // --- small XML helpers ---

        private static string ParseTitleText(XElement titleEl, WorkbookModel workbook)
        {
            var textSource = titleEl.Element(C + "tx");
            var titleFormula = TextOf(FirstDescendant(textSource != null ? textSource.Element(C + "strRef") : null, C + "f"));
            if (!string.IsNullOrEmpty(titleFormula))
            {
                var resolvedTitle = ChartWorkbookDataResolver.ResolveSingleString(workbook, titleFormula);
                if (!string.IsNullOrEmpty(resolvedTitle))
                {
                    return resolvedTitle;
                }
            }

            var cachedValue = FirstDescendant(textSource, C + "v");
            if (cachedValue != null)
            {
                var cachedText = TextOf(cachedValue);
                if (!string.IsNullOrEmpty(cachedText))
                {
                    return cachedText;
                }
            }

            return ConcatText(titleEl);
        }

        private static string ParseSeriesName(XElement textSource, WorkbookModel workbook)
        {
            var nameFormula = TextOf(FirstDescendant(textSource != null ? textSource.Element(C + "strRef") : null, C + "f"));
            if (!string.IsNullOrEmpty(nameFormula))
            {
                var resolvedName = ChartWorkbookDataResolver.ResolveSingleString(workbook, nameFormula);
                if (!string.IsNullOrEmpty(resolvedName))
                {
                    return resolvedName;
                }
            }

            var cachedValue = FirstDescendant(textSource, C + "v");
            if (cachedValue != null)
            {
                var cachedText = TextOf(cachedValue);
                if (!string.IsNullOrEmpty(cachedText))
                {
                    return cachedText;
                }
            }

            return ConcatText(textSource);
        }

        private static string ConcatText(XElement element)
        {
            if (element == null)
            {
                return string.Empty;
            }

            var sb = new System.Text.StringBuilder();
            foreach (var t in element.Descendants(A + "t"))
            {
                sb.Append(t.Value);
            }
            return sb.ToString();
        }

        private static XElement FirstElement(XElement parent, XName name)
        {
            foreach (var e in parent.Elements(name)) return e;
            return null;
        }

        private static XElement FirstDescendant(XElement parent, XName name)
        {
            if (parent == null) return null;
            foreach (var e in parent.Descendants(name)) return e;
            return null;
        }

        private static XElement Descendant(XElement parent, XName name)
        {
            return FirstDescendant(parent, name);
        }

        private static XElement FirstNonNull(XElement a, XElement b)
        {
            return a ?? b;
        }

        private static string TextOf(XElement e)
        {
            return e != null ? e.Value : string.Empty;
        }

        private static string AttrOf(XElement e, string name)
        {
            if (e == null) return null;
            var a = e.Attribute(name);
            return a != null ? a.Value : null;
        }

        private static bool TryDouble(string s, out double value)
        {
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryHex(string hex, out SKColor color)
        {
            color = SKColors.Gray;
            if (string.IsNullOrEmpty(hex) || hex.Length < 6) return false;
            int offset = hex.Length == 8 ? 2 : 0;
            int r, g, b;
            if (int.TryParse(hex.Substring(offset, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r)
                && int.TryParse(hex.Substring(offset + 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g)
                && int.TryParse(hex.Substring(offset + 4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b))
            {
                color = new SKColor((byte)r, (byte)g, (byte)b);
                return true;
            }
            return false;
        }
    }
}
