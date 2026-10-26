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
        Scatter,
    }

    internal sealed class ChartSeries
    {
        public string Name = string.Empty;
        public SKColor Color = SKColors.Gray;
        public Dictionary<int, SKColor> PointColors = new Dictionary<int, SKColor>();
        public SKColor LineColor = SKColors.Gray;
        public double LineWidthPt = 2.25d;
        public bool HasVisibleLine = true;
        public bool HasBarBorder;
        public bool SmoothLine;
        public List<double?> Values = new List<double?>();
        public List<double?> XValues = new List<double?>();
        public List<double?> BubbleSizes = new List<double?>();
        public bool Bubble3D;
        public bool MarkerVisible;
        public string MarkerSymbol = "none";
        public int MarkerSize = 5;
        public SKColor MarkerFillColor = SKColors.Gray;
        public SKColor MarkerStrokeColor = SKColors.Gray;
        public double MarkerStrokeWidthPt = 0.75d;
        public string PatternPreset;
        public SKColor PatternForegroundColor = SKColors.Transparent;
        public SKColor PatternBackgroundColor = SKColors.Transparent;
        public bool ShowValueDataLabels;
        public string DataLabelPosition;
        public double? DataLabelSizePt;
        public string ValueFormatCode;
        public SKColor DataLabelColor = new SKColor(0x40, 0x40, 0x40);
        public bool HasErrorBars;
        public string ErrorBarType;
        public string ErrorValueType;
        public bool ErrorBarsHaveEndCaps = true;
        public SKColor ErrorBarColor = new SKColor(0x59, 0x59, 0x59);
        public double ErrorBarWidthPt = 0.75d;
    }

    /// <summary>
    /// A chart's definition reduced to what the renderer needs, parsed from the cached values in the
    /// chart XML (no formula engine required).
    /// </summary>
    internal sealed class ParsedChart
    {
        public ChartKind Kind = ChartKind.Unsupported;
        public CultureInfo Culture;
        public double? TextSizePt;
        public string Title;                 // null = no title band
        public FontValue TitleFont = new FontValue { Name = "Calibri", Size = 14d, Bold = true };
        public SKColor TitleColor = new SKColor(0x59, 0x59, 0x59);
        public SKColor? TitleTextOutlineColor;
        public bool TitleAllCaps;
        public bool HasTitleBorder;
        public SKColor TitleBorderColor = SKColors.Transparent;
        public double TitleBorderWidthPt = 0.75d;
        public SKColor AxisLabelColor = new SKColor(0x59, 0x59, 0x59);
        public SKColor? AxisTextOutlineColor;
        public SKColor LegendTextColor = new SKColor(0x59, 0x59, 0x59);
        public SKColor? LegendTextOutlineColor;
        public SKColor SeriesAxisLabelColor = new SKColor(0x59, 0x59, 0x59);
        public string CategoryAxisTitle;
        public string ValueAxisTitle;
        public bool HasChartAreaFill;
        public SKColor ChartAreaFillColor = SKColors.White;
        public string ChartAreaPatternPreset;
        public SKColor ChartAreaPatternForegroundColor = SKColors.Transparent;
        public SKColor ChartAreaPatternBackgroundColor = SKColors.White;
        public bool HasChartAreaBorder;
        public SKColor ChartAreaBorderColor = new SKColor(0xD9, 0xD9, 0xD9);
        public double ChartAreaBorderWidthPt = 1d;
        public bool HasChartAreaGlow;
        public SKColor ChartAreaGlowColor;
        public double ChartAreaGlowRadiusPt;
        public double ChartAreaSoftEdgeRadiusPt;
        public bool HasChartAreaGradient;
        public List<SKColor> ChartAreaGradientColors = new List<SKColor>();
        public List<float> ChartAreaGradientPositions = new List<float>();
        public double ChartAreaGradientAngleDeg;
        public string LegendPosition;        // "b","t","l","r","tr", or null
        public bool HasManualLegend;
        public double LegendX;
        public double LegendY;
        public double LegendW;
        public double LegendH;
        public bool HasValueGridlines;
        public SKColor ValueGridlineColor = new SKColor(0xD9, 0xD9, 0xD9);
        public double ValueGridlineWidthPt = 1d;
        public bool HasMinorValueGridlines;
        public bool HasValueAxisLine;
        public bool ValueAxisOnRight;
        public SKColor ValueAxisLineColor = new SKColor(0x59, 0x59, 0x59);
        public double ValueAxisLineWidthPt = 0.75d;
        public string ValueAxisMajorTickMark;
        public string ValueFormatCode;
        public bool CrossBetween = true;  // category points sit at band centers (Excel default)
        public bool IsStacked;
        public bool IsPercentStacked;
        public double? AxisMin;              // explicit scaling, else auto
        public double? AxisMax;
        public double? MajorUnit;
        public double? MinorUnit;
        public double? ScatterXMin;
        public double? ScatterXMax;
        public double? ScatterXMajorUnit;
        public string ScatterXFormatCode;
        public bool ScatterVaryColors;
        public bool IsBubble;
        public double BubbleScalePercent = 100d;
        public bool ShowNegativeBubbles;
        public bool ScatterValueAxisDeleted;
        public bool ScatterCategoryAxisDeleted;
        public bool ScatterTitleOverlay;
        public bool RoundedCorners;
        public SKColor MinorValueGridlineColor = new SKColor(0xEDEDED);
        public double MinorValueGridlineWidthPt = 1d;
        public bool HasCategoryGridlines;
        public SKColor CategoryGridlineColor = new SKColor(0xD9, 0xD9, 0xD9);
        public double CategoryGridlineWidthPt = 1d;
        public bool HasPlotAreaBorder;
        public SKColor PlotAreaBorderColor = SKColors.Transparent;
        public double PlotAreaBorderWidthPt = 1d;
        public bool HasDataTable;
        public bool DataTableShowHorizontalBorders;
        public bool DataTableShowVerticalBorders;
        public bool DataTableShowOutline;
        public bool DataTableShowKeys;
        public SKColor DataTableLineColor = new SKColor(0xD9, 0xD9, 0xD9);
        public double DataTableLineWidthPt = 0.75d;
        public SKColor DataTableTextColor = new SKColor(0x59, 0x59, 0x59);
        public List<ChartSeries> Series = new List<ChartSeries>();
        public List<string> Categories = new List<string>();
        public List<double?> CategoryValues = new List<double?>();
        public List<SKColor> PiePointColors = new List<SKColor>();
        public List<SKColor> PiePointLineColors = new List<SKColor>();
        public List<double> PiePointLineWidthsPt = new List<double>();
        public bool IsDoughnut;
        public double DoughnutHoleSizePercent = 50d;
        public double DoughnutFirstSliceAngleDeg;
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
        public string CategoryAxisMajorTickMark;
        public bool HasCategoryAxisLine = true;
        public SKColor CategoryAxisLineColor = new SKColor(0xD9, 0xD9, 0xD9);
        public double CategoryAxisLineWidthPt = 1d;
        public bool CategoryAxisAutomaticTextRotation;
        public bool CategoryAxisMinMax = true;
        public string RadarStyle = "standard";
        public double GapWidthPercent = 150d;
        public double OverlapPercent;
        public string BarShape = "box";
        public bool IsLine3D;
        public bool IsArea3D;
        public bool IsBar3D;
        public bool IsPie3D;
        public double ValueDisplayUnitDivisor = 1d;
        public string ValueDisplayUnitLabel;
        public bool HasManualValueDisplayUnitLabel;
        public double ValueDisplayUnitLabelX;
        public double ValueDisplayUnitLabelY;
        public double RotationXDeg = 15d;
        public double RotationYDeg = 20d;
        public double PerspectiveDeg = 30d;
        public double DepthPercent = 100d;
        public double GapDepthPercent = 150d;
        public bool RightAngleAxes = true;
        public bool HasVisibleSeriesAxis;
        public byte[] SideWallImageData;
        public byte[] BackWallImageData;

        public bool HasManualTitle;
        public double TitleX;
        public double TitleY;

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
            result.IsBubble = typeEl.Name == C + "bubbleChart";
            if (result.IsBubble)
            {
                double bubbleScale;
                if (TryDouble(AttrOf(typeEl.Element(C + "bubbleScale"), "val"), out bubbleScale)
                    && bubbleScale > 0d)
                {
                    result.BubbleScalePercent = bubbleScale;
                }
                result.ShowNegativeBubbles = string.Equals(
                    AttrOf(typeEl.Element(C + "showNegBubbles"), "val"), "1", StringComparison.Ordinal);
            }
            result.RoundedCorners = string.Equals(AttrOf(doc.Root.Element(C + "roundedCorners"), "val"), "1", StringComparison.Ordinal);
            if (kind == ChartKind.Scatter)
            {
                result.ScatterVaryColors = string.Equals(AttrOf(typeEl.Element(C + "varyColors"), "val"), "1", StringComparison.Ordinal);
            }
            var scatterStyle = AttrOf(typeEl.Element(C + "scatterStyle"), "val");
            var smoothScatterStyle = string.Equals(scatterStyle, "smooth", StringComparison.Ordinal)
                || string.Equals(scatterStyle, "smoothMarker", StringComparison.Ordinal);
            result.Culture = ResolveChartCulture(doc.Root, culture);
            ChartExplicitFormatting.ApplyTextDefaults(doc.Root, colors, result);
            result.Title = ParseTitle(chart, colors, result, chartModel, workbook);
            ParseChartAreaStyle(doc.Root, colors, result);
            ParseLegend(chart, colors, result, chartModel);
            ParseManualPlot(plotArea, result);
            ParseThreeDimensionalSettings(chart, typeEl, result, chartModel);
            ParseDoughnutSettings(typeEl, result);
            ParseOfPieSettings(typeEl, colors, result);
            ParseGrouping(typeEl, result);
            ParseRadarSettings(typeEl, result);
            ParseBarSeriesLayout(typeEl, result);
            ParseValueAxis(plotArea, colors, result, chartModel, workbook);
            ParseCategoryAxis(plotArea, colors, result, chartModel, workbook);
            ParseSeriesAxis(plotArea, colors, result, chartModel);
            ParsePlotAreaStyle(plotArea, colors, result);
            ParseDataTable(plotArea, colors, result, chartModel);
            ChartLegacyGridStyle.Apply(doc.Root, plotArea, chartModel, result);

            var chartDataLabels = typeEl.Element(C + "dLbls");
            var seriesIndex = 0;
            foreach (var ser in typeEl.Elements(C + "ser"))
            {
                var seriesDataLabels = ser.Element(C + "dLbls") ?? chartDataLabels;
                var valueHolder = ser.Element(C + "val") ?? ser.Element(C + "yVal");
                var series = new ChartSeries
                {
                    Name = ParseSeriesName(ser.Element(C + "tx"), workbook),
                    Color = ParseSeriesColor(ser, colors, seriesIndex),
                    LineWidthPt = ParseLineWidthPt(ser),
                    SmoothLine = kind == ChartKind.Scatter && (ser.Element(C + "smooth") == null
                        ? smoothScatterStyle
                        : string.Equals(AttrOf(ser.Element(C + "smooth"), "val"), "1", StringComparison.Ordinal)),
                    Values = ParseNumericData(valueHolder, workbook, dateSystem),
                    XValues = ParseNumericData(ser.Element(C + "xVal"), workbook, dateSystem),
                    BubbleSizes = ParseNumericData(ser.Element(C + "bubbleSize"), workbook, dateSystem),
                    Bubble3D = string.Equals(AttrOf(ser.Element(C + "bubble3D"), "val"),
                        "1", StringComparison.Ordinal),
                    ShowValueDataLabels = string.Equals(
                        AttrOf(seriesDataLabels != null ? seriesDataLabels.Element(C + "showVal") : null, "val"),
                        "1",
                        StringComparison.Ordinal),
                    DataLabelPosition = AttrOf(seriesDataLabels != null ? seriesDataLabels.Element(C + "dLblPos") : null, "val"),
                    ValueFormatCode = TextOf(FirstDescendant(valueHolder, C + "formatCode")),
                };
                ChartExplicitFormatting.ApplySeries(ser, typeEl, colors, series, seriesIndex);
                ParseSeriesLineStyle(ser, colors, series);
                if ((kind == ChartKind.Column || kind == ChartKind.Bar) && !result.IsBar3D)
                    ChartBarBorder.Apply(ser, series);
                ParseSeriesPatternFill(ser, colors, series);
                ParseMarker(ser, colors, series);
                ParseErrorBars(ser, colors, series);
                SKColor dataLabelColor;
                if (TryResolveTextColor(seriesDataLabels != null ? seriesDataLabels.Element(C + "txPr") : null,
                    colors, chartModel, out dataLabelColor))
                {
                    series.DataLabelColor = dataLabelColor;
                }
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
                PieChartPointMetadata.Apply(firstSer, colors, result);
                result.PiePointColors = ParsePiePointColors(firstSer, colors, result.Categories.Count);
                if (string.IsNullOrEmpty(result.ValueFormatCode))
                {
                    result.ValueFormatCode = TextOf(FirstDescendant(firstSer.Element(C + "val"), C + "formatCode"));
                }
            }

            return result;
        }

        private static void ParseDataTable(XElement plotArea, RenderColor colors, ParsedChart result, ChartModel chartModel)
        {
            var dataTable = plotArea != null ? plotArea.Element(C + "dTable") : null;
            if (dataTable == null || result == null)
            {
                return;
            }

            result.HasDataTable = true;
            result.DataTableShowHorizontalBorders = IsEnabled(dataTable.Element(C + "showHorzBorder"));
            result.DataTableShowVerticalBorders = IsEnabled(dataTable.Element(C + "showVertBorder"));
            result.DataTableShowOutline = IsEnabled(dataTable.Element(C + "showOutline"));
            result.DataTableShowKeys = IsEnabled(dataTable.Element(C + "showKeys"));

            var line = dataTable.Element(C + "spPr") != null
                ? dataTable.Element(C + "spPr").Element(A + "ln")
                : null;
            bool hasVisibleLine;
            SKColor lineColor;
            double lineWidth;
            ParseLineStyle(line, colors, out hasVisibleLine, out lineColor, out lineWidth);
            if (hasVisibleLine)
            {
                result.DataTableLineColor = lineColor;
                result.DataTableLineWidthPt = lineWidth;
            }

            SKColor textColor;
            if (TryResolveTextColor(dataTable.Element(C + "txPr"), colors, chartModel, out textColor))
            {
                result.DataTableTextColor = textColor;
            }
        }

        private static void ParseErrorBars(XElement seriesElement, RenderColor colors, ChartSeries series)
        {
            var errorBars = seriesElement != null ? seriesElement.Element(C + "errBars") : null;
            if (errorBars == null || series == null)
            {
                return;
            }

            series.HasErrorBars = true;
            series.ErrorBarType = AttrOf(errorBars.Element(C + "errBarType"), "val");
            series.ErrorValueType = AttrOf(errorBars.Element(C + "errValType"), "val");
            series.ErrorBarsHaveEndCaps = !string.Equals(
                AttrOf(errorBars.Element(C + "noEndCap"), "val"), "1", StringComparison.Ordinal);

            var line = errorBars.Element(C + "spPr") != null
                ? errorBars.Element(C + "spPr").Element(A + "ln")
                : null;
            bool hasVisibleLine;
            SKColor lineColor;
            double lineWidth;
            ParseLineStyle(line, colors, out hasVisibleLine, out lineColor, out lineWidth);
            if (hasVisibleLine)
            {
                series.ErrorBarColor = lineColor;
                series.ErrorBarWidthPt = lineWidth;
            }
        }

        private static bool IsEnabled(XElement element)
        {
            return element != null
                && !string.Equals(AttrOf(element, "val"), "0", StringComparison.Ordinal);
        }

        private static void ApplyFallbackTitle(ParsedChart chart)
        {
            if (chart == null)
            {
                return;
            }

            if ((chart.Kind == ChartKind.Pie && !chart.IsDoughnut
                    || chart.Kind == ChartKind.Scatter && chart.Series != null && chart.Series.Count == 1)
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

            var scatter = plotArea.Element(C + "scatterChart");
            if (scatter != null) { kind = ChartKind.Scatter; return scatter; }

            var bubble = plotArea.Element(C + "bubbleChart");
            if (bubble != null) { kind = ChartKind.Scatter; return bubble; }

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

            var doughnut = plotArea.Element(C + "doughnutChart");
            if (doughnut != null) { kind = ChartKind.Pie; return doughnut; }

            var ofPie = plotArea.Element(C + "ofPieChart");
            if (ofPie != null) { kind = ChartKind.Pie; return ofPie; }

            kind = ChartKind.Unsupported;
            return null;
        }

        private static void ParseDoughnutSettings(XElement typeEl, ParsedChart result)
        {
            if (typeEl == null || typeEl.Name != C + "doughnutChart")
            {
                return;
            }

            result.IsDoughnut = true;
            double value;
            if (TryDouble(AttrOf(typeEl.Element(C + "holeSize"), "val"), out value))
            {
                result.DoughnutHoleSizePercent = Math.Max(10d, Math.Min(90d, value));
            }
            if (TryDouble(AttrOf(typeEl.Element(C + "firstSliceAng"), "val"), out value))
            {
                result.DoughnutFirstSliceAngleDeg = value;
            }
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

            ParseManualTitleLayout(titleEl, result);
            result.ScatterTitleOverlay = result.Kind == ChartKind.Scatter
                && string.Equals(AttrOf(titleEl.Element(C + "overlay"), "val"), "1", StringComparison.Ordinal);
            ParseTitleStyle(titleEl, colors, result, chartModel);
            var text = ParseTitleText(titleEl, workbook);
            if (!string.IsNullOrEmpty(text))
            {
                return text;
            }

            return "Chart Title";
        }

        private static void ParseManualTitleLayout(XElement titleEl, ParsedChart result)
        {
            var manual = titleEl.Element(C + "layout") != null
                ? titleEl.Element(C + "layout").Element(C + "manualLayout")
                : null;
            if (manual == null)
            {
                return;
            }

            double x;
            double y;
            if (TryDouble(AttrOf(manual.Element(C + "x"), "val"), out x)
                && TryDouble(AttrOf(manual.Element(C + "y"), "val"), out y))
            {
                result.HasManualTitle = true;
                result.TitleX = x;
                result.TitleY = y;
            }
        }

        private static void ParseTitleStyle(XElement titleEl, RenderColor colors, ParsedChart result, ChartModel chartModel)
        {
            if (titleEl == null || result == null)
            {
                return;
            }

            var titleTextProperties = ResolveTextProperties(titleEl);

            var defaultRunProperties = titleTextProperties != null
                ? FirstDescendant(titleTextProperties, A + "defRPr")
                : null;
            result.TitleAllCaps = string.Equals(AttrOf(defaultRunProperties, "cap"), "all", StringComparison.Ordinal);

            SKColor titleColor;
            if (TryResolveTextColor(titleTextProperties, colors, chartModel, out titleColor))
            {
                result.TitleColor = titleColor;
            }
            if (TryResolveTextOutlineColor(titleTextProperties, colors, out titleColor))
            {
                result.TitleTextOutlineColor = titleColor;
            }

            FontValue titleFont;
            if (TryResolveTextFont(titleTextProperties, colors, out titleFont))
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
                var layout = legend.Element(C + "layout");
                var manual = layout != null ? layout.Element(C + "manualLayout") : null;
                double x;
                double y;
                double w;
                double h;
                if (manual != null
                    && TryDouble(AttrOf(manual.Element(C + "x"), "val"), out x)
                    && TryDouble(AttrOf(manual.Element(C + "y"), "val"), out y)
                    && TryDouble(AttrOf(manual.Element(C + "w"), "val"), out w)
                    && TryDouble(AttrOf(manual.Element(C + "h"), "val"), out h))
                {
                    result.HasManualLegend = true;
                    result.LegendX = x;
                    result.LegendY = y;
                    result.LegendW = w;
                    result.LegendH = h;
                }

                SKColor legendTextColor;
                if (TryResolveTextColor(legend.Element(C + "txPr"), colors, chartModel, out legendTextColor))
                {
                    result.LegendTextColor = legendTextColor;
                }
                if (TryResolveTextOutlineColor(legend.Element(C + "txPr"), colors, out legendTextColor))
                {
                    result.LegendTextOutlineColor = legendTextColor;
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

            var shape = AttrOf(typeEl.Element(C + "shape"), "val");
            if (!string.IsNullOrEmpty(shape))
            {
                result.BarShape = shape;
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

        private static void ParseThreeDimensionalSettings(XElement chart, XElement typeEl, ParsedChart result, ChartModel chartModel)
        {
            if (chart == null || typeEl == null || result == null)
            {
                return;
            }

            var isLine3D = typeEl.Name == C + "line3DChart";
            var isArea3D = typeEl.Name == C + "area3DChart";
            var isBar3D = typeEl.Name == C + "bar3DChart";
            var isPie3D = typeEl.Name == C + "pie3DChart";
            if (!isLine3D && !isArea3D && !isBar3D && !isPie3D)
            {
                return;
            }

            result.IsLine3D = isLine3D;
            result.IsArea3D = isArea3D;
            result.IsBar3D = isBar3D;
            result.IsPie3D = isPie3D;
            result.SideWallImageData = ResolveWallImageData(chart.Element(C + "sideWall"), chartModel);
            result.BackWallImageData = ResolveWallImageData(chart.Element(C + "backWall"), chartModel);
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

            if (TryDouble(AttrOf(view3D != null ? view3D.Element(C + "depthPercent") : null, "val"), out value) && value > 0d)
            {
                result.DepthPercent = value;
            }

            var rightAngleAxes = AttrOf(view3D != null ? view3D.Element(C + "rAngAx") : null, "val");
            if (!string.IsNullOrEmpty(rightAngleAxes))
            {
                result.RightAngleAxes = !string.Equals(rightAngleAxes, "0", StringComparison.Ordinal)
                    && !string.Equals(rightAngleAxes, "false", StringComparison.OrdinalIgnoreCase);
            }

            if (TryDouble(AttrOf(typeEl.Element(C + "gapDepth"), "val"), out value) && value >= 0d)
            {
                result.GapDepthPercent = value;
            }
        }

        private static byte[] ResolveWallImageData(XElement wall, ChartModel chartModel)
        {
            if (wall == null || chartModel == null)
            {
                return null;
            }

            var shapeProperties = wall.Element(C + "spPr");
            var blipFill = shapeProperties != null ? shapeProperties.Element(A + "blipFill") : null;
            var blip = blipFill != null ? blipFill.Element(A + "blip") : null;
            var relationshipId = blip != null
                ? (string)blip.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))
                : null;
            if (string.IsNullOrEmpty(relationshipId) || blipFill.Element(A + "tile") == null)
            {
                return null;
            }

            for (var i = 0; i < chartModel.CompanionFiles.Count; i++)
            {
                var companion = chartModel.CompanionFiles[i];
                if (string.Equals(companion.RelationshipId, relationshipId, StringComparison.Ordinal)
                    && companion.BinaryContent != null
                    && companion.BinaryContent.Length > 0)
                {
                    return companion.BinaryContent;
                }
            }

            return null;
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

        private static void ParseValueAxis(XElement plotArea, RenderColor colors, ParsedChart result, ChartModel chartModel, WorkbookModel workbook)
        {
            var valAx = result.Kind == ChartKind.Scatter
                ? FindValueAxisAt(plotArea, "l", "r")
                : FirstElement(plotArea, C + "valAx");
            if (valAx == null)
            {
                return;
            }

            result.HasValueGridlines = valAx.Element(C + "majorGridlines") != null;
            result.HasMinorValueGridlines = valAx.Element(C + "minorGridlines") != null;
            result.ValueAxisMajorTickMark = AttrOf(valAx.Element(C + "majorTickMark"), "val");
            result.ValueAxisOnRight = string.Equals(AttrOf(valAx.Element(C + "axPos"), "val"), "r", StringComparison.Ordinal);
            result.ValueFormatCode = AttrOf(valAx.Element(C + "numFmt"), "formatCode");
            result.ValueAxisTitle = ParseAxisTitle(valAx, workbook);
            result.ScatterValueAxisDeleted = result.Kind == ChartKind.Scatter
                && string.Equals(AttrOf(valAx.Element(C + "delete"), "val"), "1", StringComparison.Ordinal);
            var displayUnits = valAx.Element(C + "dispUnits");
            var displayUnit = AttrOf(displayUnits != null
                ? displayUnits.Element(C + "builtInUnit")
                : null, "val");
            if (string.Equals(displayUnit, "hundreds", StringComparison.Ordinal))
            {
                result.ValueDisplayUnitDivisor = 100d;
                result.ValueDisplayUnitLabel = "Hundreds";
            }
            var displayUnitLabel = displayUnits != null ? displayUnits.Element(C + "dispUnitsLbl") : null;
            var displayUnitLayout = displayUnitLabel != null ? displayUnitLabel.Element(C + "layout") : null;
            var displayUnitManual = displayUnitLayout != null ? displayUnitLayout.Element(C + "manualLayout") : null;
            double displayUnitX;
            double displayUnitY;
            if (displayUnitManual != null
                && TryDouble(AttrOf(displayUnitManual.Element(C + "x"), "val"), out displayUnitX)
                && TryDouble(AttrOf(displayUnitManual.Element(C + "y"), "val"), out displayUnitY))
            {
                result.HasManualValueDisplayUnitLabel = true;
                result.ValueDisplayUnitLabelX = displayUnitX;
                result.ValueDisplayUnitLabelY = displayUnitY;
            }
            result.CrossBetween = !string.Equals(AttrOf(valAx.Element(C + "crossBetween"), "val"), "midCat", StringComparison.Ordinal);

            var shapeProperties = valAx.Element(C + "spPr");
            var axisLine = shapeProperties != null ? shapeProperties.Element(A + "ln") : null;
            ParseLineStyle(axisLine, colors, out result.HasValueAxisLine, out result.ValueAxisLineColor, out result.ValueAxisLineWidthPt);
            if (axisLine != null && axisLine.Attribute("w") == null)
            {
                result.ValueAxisLineWidthPt = 0.75d;
            }
            ParseValueGridlineStyle(valAx.Element(C + "majorGridlines"), colors, result, false);
            ParseValueGridlineStyle(valAx.Element(C + "minorGridlines"), colors, result, true);
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
            if (TryResolveTextOutlineColor(valAx.Element(C + "txPr"), colors, out axisLabelColor))
            {
                result.AxisTextOutlineColor = axisLabelColor;
            }
        }

        private static void ParseCategoryAxis(XElement plotArea, RenderColor colors, ParsedChart result, ChartModel chartModel, WorkbookModel workbook)
        {
            var dateAxis = FirstElement(plotArea, C + "dateAx");
            var categoryAxis = FirstElement(plotArea, C + "catAx");
            var scatterAxis = result.Kind == ChartKind.Scatter
                ? FindValueAxisAt(plotArea, "b", "t")
                : null;
            var axisElement = dateAxis ?? categoryAxis ?? scatterAxis;
            if (axisElement == null)
            {
                return;
            }

            result.HasDateCategoryAxis = dateAxis != null;
            result.CategoryFormatCode = AttrOf(axisElement.Element(C + "numFmt"), "formatCode");
            if (scatterAxis != null)
            {
                result.ScatterXFormatCode = result.CategoryFormatCode;
                var scatterScaling = scatterAxis.Element(C + "scaling");
                double scatterValue;
                if (scatterScaling != null)
                {
                    if (TryDouble(AttrOf(scatterScaling.Element(C + "min"), "val"), out scatterValue)) result.ScatterXMin = scatterValue;
                    if (TryDouble(AttrOf(scatterScaling.Element(C + "max"), "val"), out scatterValue)) result.ScatterXMax = scatterValue;
                }
                if (TryDouble(AttrOf(scatterAxis.Element(C + "majorUnit"), "val"), out scatterValue)) result.ScatterXMajorUnit = scatterValue;
            }
            result.CategoryAxisTitle = ParseAxisTitle(axisElement, workbook);
            result.ScatterCategoryAxisDeleted = result.Kind == ChartKind.Scatter
                && string.Equals(AttrOf(axisElement.Element(C + "delete"), "val"), "1", StringComparison.Ordinal);
            result.CategoryBaseTimeUnit = AttrOf(axisElement.Element(C + "baseTimeUnit"), "val");
            result.CategoryAxisMajorTickMark = AttrOf(axisElement.Element(C + "majorTickMark"), "val");
            var axisShape = axisElement.Element(C + "spPr");
            var axisLine = axisShape != null ? axisShape.Element(A + "ln") : null;
            if (axisLine != null)
            {
                ParseLineStyle(axisLine, colors, out result.HasCategoryAxisLine,
                    out result.CategoryAxisLineColor, out result.CategoryAxisLineWidthPt);
            }
            var scaling = axisElement.Element(C + "scaling");
            result.CategoryAxisMinMax = !string.Equals(AttrOf(scaling != null ? scaling.Element(C + "orientation") : null, "val"), "maxMin", StringComparison.Ordinal);
            ParseGridlineStyle(axisElement.Element(C + "majorGridlines"), colors, result);

            var bodyPr = axisElement.Element(C + "txPr") != null ? axisElement.Element(C + "txPr").Element(A + "bodyPr") : null;
            double rotation;
            if (TryDouble(AttrOf(bodyPr, "rot"), out rotation))
            {
                // Excel's automatic orientation marker is not a literal -1000 degree angle.
                result.CategoryAxisAutomaticTextRotation = rotation == -60000000d;
                result.CategoryAxisTextRotationDeg = Math.Max(-90d, Math.Min(90d, rotation / 60000d));
            }

            SKColor axisLabelColor;
            if (TryResolveTextColor(axisElement.Element(C + "txPr"), colors, chartModel, out axisLabelColor))
            {
                result.AxisLabelColor = axisLabelColor;
            }
            if (TryResolveTextOutlineColor(axisElement.Element(C + "txPr"), colors, out axisLabelColor))
            {
                result.AxisTextOutlineColor = axisLabelColor;
            }
        }

        private static XElement FindValueAxisAt(XElement plotArea, string firstPosition, string secondPosition)
        {
            if (plotArea == null)
            {
                return null;
            }

            foreach (var axis in plotArea.Elements(C + "valAx"))
            {
                var position = AttrOf(axis.Element(C + "axPos"), "val");
                if (string.Equals(position, firstPosition, StringComparison.Ordinal)
                    || string.Equals(position, secondPosition, StringComparison.Ordinal))
                {
                    return axis;
                }
            }

            return null;
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
            if (result.Kind == ChartKind.Scatter && line != null && AttrOf(line, "w") == null)
            {
                result.PlotAreaBorderWidthPt = 0.75d;
            }
        }

        private static void ParseChartAreaStyle(XElement chartSpace, RenderColor colors, ParsedChart result)
        {
            if (chartSpace == null || colors == null || result == null)
            {
                return;
            }

            var shapeProperties = chartSpace.Element(C + "spPr");
            var borderLine = shapeProperties != null ? shapeProperties.Element(A + "ln") : null;
            ParseLineStyle(borderLine, colors, out result.HasChartAreaBorder,
                out result.ChartAreaBorderColor, out result.ChartAreaBorderWidthPt);
            var patternFill = shapeProperties != null ? shapeProperties.Element(A + "pattFill") : null;
            if (patternFill != null)
            {
                result.ChartAreaPatternPreset = AttrOf(patternFill, "prst");
                SKColor patternColor;
                if (TryDrawingColor(FillChild(patternFill.Element(A + "fgClr")), colors, out patternColor))
                {
                    result.ChartAreaPatternForegroundColor = patternColor;
                }
                if (TryDrawingColor(FillChild(patternFill.Element(A + "bgClr")), colors, out patternColor))
                {
                    result.ChartAreaPatternBackgroundColor = patternColor;
                }
            }

            var effectList = shapeProperties != null ? shapeProperties.Element(A + "effectLst") : null;
            var glow = effectList != null ? effectList.Element(A + "glow") : null;
            double glowRadius;
            SKColor glowColor;
            if (TryDouble(AttrOf(glow, "rad"), out glowRadius)
                && glowRadius > 0d
                && TryDrawingColor(FillChild(glow), colors, out glowColor))
            {
                result.HasChartAreaGlow = true;
                result.ChartAreaGlowRadiusPt = glowRadius / 12700d;
                result.ChartAreaGlowColor = glowColor;
            }
            var softEdge = effectList != null ? effectList.Element(A + "softEdge") : null;
            double softEdgeRadius;
            if (TryDouble(AttrOf(softEdge, "rad"), out softEdgeRadius) && softEdgeRadius > 0d)
            {
                result.ChartAreaSoftEdgeRadiusPt = softEdgeRadius / 12700d;
            }
            var gradientFill = shapeProperties != null ? shapeProperties.Element(A + "gradFill") : null;
            if (gradientFill != null)
            {
                var stops = gradientFill.Element(A + "gsLst");
                var parsedStops = new List<Tuple<float, SKColor>>();
                if (stops != null)
                {
                    foreach (var stop in stops.Elements(A + "gs"))
                    {
                        double position;
                        SKColor stopColor;
                        if (TryDouble(AttrOf(stop, "pos"), out position)
                            && TryDrawingColor(FillChild(stop), colors, out stopColor))
                        {
                            parsedStops.Add(Tuple.Create((float)(position / 100000d), stopColor));
                        }
                    }
                }

                parsedStops.Sort(delegate(Tuple<float, SKColor> left, Tuple<float, SKColor> right)
                {
                    return left.Item1.CompareTo(right.Item1);
                });
                for (var i = 0; i < parsedStops.Count; i++)
                {
                    result.ChartAreaGradientPositions.Add(parsedStops[i].Item1);
                    result.ChartAreaGradientColors.Add(parsedStops[i].Item2);
                }

                double angle;
                var linear = gradientFill.Element(A + "lin");
                if (TryDouble(AttrOf(linear, "ang"), out angle))
                {
                    result.ChartAreaGradientAngleDeg = angle / 60000d;
                }

                result.HasChartAreaGradient = result.ChartAreaGradientColors.Count >= 2;
                if (result.HasChartAreaGradient)
                {
                    return;
                }
            }

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

        private static void ParseValueGridlineStyle(XElement gridlines, RenderColor colors, ParsedChart result, bool minor)
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
            if (minor)
            {
                result.HasMinorValueGridlines = hasVisibleLine || line == null;
                if (hasVisibleLine)
                {
                    result.MinorValueGridlineColor = color;
                    result.MinorValueGridlineWidthPt = widthPt;
                }
            }
            else
            {
                result.HasValueGridlines = hasVisibleLine || line == null;
                if (hasVisibleLine)
                {
                    result.ValueGridlineColor = color;
                    result.ValueGridlineWidthPt = widthPt;
                }
            }
        }

        internal static void ParseLineStyle(XElement line, RenderColor colors, out bool hasVisibleLine, out SKColor color, out double widthPt)
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

        private static void ParseSeriesPatternFill(XElement ser, RenderColor colors, ChartSeries series)
        {
            var shapeProperties = ser != null ? ser.Element(C + "spPr") : null;
            var patternFill = shapeProperties != null ? shapeProperties.Element(A + "pattFill") : null;
            if (patternFill == null || series == null)
            {
                return;
            }

            series.PatternPreset = AttrOf(patternFill, "prst");
            SKColor parsedColor;
            var foreground = FillChild(patternFill.Element(A + "fgClr"));
            if (TryDrawingColor(foreground, colors, out parsedColor))
            {
                series.PatternForegroundColor = parsedColor;
                series.Color = parsedColor;
            }

            var background = FillChild(patternFill.Element(A + "bgClr"));
            if (TryDrawingColor(background, colors, out parsedColor))
            {
                series.PatternBackgroundColor = parsedColor;
            }
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
            var runProperties = FirstDescendant(textProperties, A + "rPr");
            if (ChartTextStyleResolver.TryResolveTextColor(runProperties, colors, chartModel, out color))
            {
                return true;
            }

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

        private static bool TryResolveTextOutlineColor(XElement textProperties, RenderColor colors, out SKColor color)
        {
            color = SKColors.Transparent;
            var runProperties = FirstDescendant(textProperties, A + "rPr");
            if (ChartTextStyleResolver.TryResolveTextOutlineColor(runProperties, colors, out color))
            {
                return true;
            }

            var defaultRunProperties = ResolveDefaultRunProperties(textProperties);
            return ChartTextStyleResolver.TryResolveTextOutlineColor(defaultRunProperties, colors, out color);
        }

        private static bool TryResolveTextFont(XElement textProperties, RenderColor colors, out FontValue font)
        {
            font = null;
            var defaultRunProperties = ResolveDefaultRunProperties(textProperties);
            if (defaultRunProperties == null)
            {
                return false;
            }

            return ChartTextStyleResolver.TryResolveFont(defaultRunProperties,
                colors.MajorFontName, colors.MinorFontName, out font);
        }

        private static XElement ResolveTextProperties(XElement owner)
        {
            if (owner == null)
            {
                return null;
            }

            var richText = owner.Element(C + "tx");
            if (richText != null)
            {
                richText = richText.Element(C + "rich");
            }

            return richText ?? owner.Element(C + "txPr");
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

        internal static bool IsDateFormat(string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return false;
            }

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

        private static string ParseAxisTitle(XElement axisElement, WorkbookModel workbook)
        {
            var title = axisElement != null ? axisElement.Element(C + "title") : null;
            if (title == null)
            {
                return null;
            }

            var text = ParseTitleText(title, workbook);
            return !string.IsNullOrEmpty(text) ? text.Trim() : "Axis Title";
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
