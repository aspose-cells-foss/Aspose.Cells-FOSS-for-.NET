using System.Globalization;
using System.Text;

namespace Aspose.Cells_FOSS
{
    internal static class ChartXmlTemplates
    {
        private const string ChartNsUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string DrawingNsUri = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        internal static string Build(ChartType type, string dataRange, string chartTitle)
        {
            switch (type)
            {
                case ChartType.Column:
                    return BuildWithAxes(dataRange, chartTitle, BuildBarInner("col", "clustered"));
                case ChartType.Bar:
                    return BuildWithAxes(dataRange, chartTitle, BuildBarInner("bar", "clustered"));
                case ChartType.Line:
                    return BuildWithAxes(dataRange, chartTitle, BuildLineInner("standard"));
                case ChartType.Area:
                    return BuildWithAxes(dataRange, chartTitle, BuildAreaInner("standard"));
                case ChartType.Pie:
                    return BuildPieStyle(dataRange, chartTitle, "pieChart");
                case ChartType.Doughnut:
                    return BuildPieStyle(dataRange, chartTitle, "doughnutChart", holeSize: 50);
                case ChartType.Scatter:
                    return BuildScatterStyle(dataRange, chartTitle);
                case ChartType.Bubble:
                    return BuildBubbleStyle(dataRange, chartTitle);
                case ChartType.Radar:
                    return BuildRadarStyle(dataRange, chartTitle);
                case ChartType.Stock:
                    return BuildStockStyle(dataRange, chartTitle);
                case ChartType.Column3D:
                    return BuildWithAxes3D(dataRange, chartTitle, BuildBar3DInner("col", "clustered"));
                case ChartType.Bar3D:
                    return BuildWithAxes3D(dataRange, chartTitle, BuildBar3DInner("bar", "clustered"));
                case ChartType.Line3D:
                    return BuildWithAxes3D(dataRange, chartTitle, BuildLine3DInner());
                case ChartType.Area3D:
                    return BuildWithAxes3D(dataRange, chartTitle, BuildArea3DInner("standard"));
                case ChartType.Pie3D:
                    return BuildPieStyle(dataRange, chartTitle, "pie3DChart");
                case ChartType.Surface3D:
                    return BuildSurface3DStyle(dataRange, chartTitle, wireframe: false);
                case ChartType.SurfaceWireframe3D:
                    return BuildSurface3DStyle(dataRange, chartTitle, wireframe: true);
                case ChartType.Contour:
                    return BuildContourStyle(dataRange, chartTitle);
                default:
                    throw new UnsupportedFeatureException(
                        "Programmatic chart creation is not supported for chart type '" + type + "'. " +
                        "This chart type can only be loaded from an existing XLSX file.");
            }
        }

        private static string Header()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                   "<c:chartSpace xmlns:c=\"" + ChartNsUri + "\"" +
                   " xmlns:a=\"" + DrawingNsUri + "\"" +
                   " xmlns:r=\"" + RelNsUri + "\">\r\n" +
                   "  <c:lang val=\"en-US\"/>\r\n";
        }

        private static string Footer()
        {
            return "</c:chartSpace>";
        }

        private static string TitleXml(string chartTitle)
        {
            if (string.IsNullOrEmpty(chartTitle))
            {
                return "    <c:autoTitleDeleted val=\"1\"/>\r\n";
            }

            var escaped = chartTitle.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
            return "    <c:title>" +
                   "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" + escaped + "</a:t></a:r></a:p></c:rich></c:tx>" +
                   "<c:overlay val=\"0\"/>" +
                   "</c:title>\r\n" +
                   "    <c:autoTitleDeleted val=\"0\"/>\r\n";
        }

        private static string SeriesWithVal(string dataRange, int idx = 0)
        {
            return "        <c:ser>" +
                   "<c:idx val=\"" + idx.ToString(CultureInfo.InvariantCulture) + "\"/>" +
                   "<c:order val=\"" + idx.ToString(CultureInfo.InvariantCulture) + "\"/>" +
                   "<c:val><c:numRef><c:f>" + dataRange + "</c:f>" +
                   "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
                   "</c:numRef></c:val>" +
                   "</c:ser>\r\n";
        }

        private static string CatAxValAx()
        {
            return "      <c:catAx>" +
                   "<c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>" +
                   "<c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/>" +
                   "</c:catAx>\r\n" +
                   "      <c:valAx>" +
                   "<c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>" +
                   "<c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/>" +
                   "</c:valAx>\r\n";
        }

        // Series axis (depth axis) required by all 3D chart types that declare three axIds.
        // deleted=true hides the axis (used by bar3D/line3D/area3D where the depth tick marks
        // are not normally shown); deleted=false makes it visible (surface/contour charts).
        private static string SerAx(bool deleted)
        {
            var deleteVal = deleted ? "1" : "0";
            return "      <c:serAx>" +
                   "<c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>" +
                   "<c:delete val=\"" + deleteVal + "\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/>" +
                   "</c:serAx>\r\n";
        }

        private static string ValAxValAx()
        {
            return "      <c:valAx>" +
                   "<c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>" +
                   "<c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/>" +
                   "</c:valAx>\r\n" +
                   "      <c:valAx>" +
                   "<c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling>" +
                   "<c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/>" +
                   "</c:valAx>\r\n";
        }

        private static string Legend()
        {
            return "    <c:legend><c:legendPos val=\"b\"/></c:legend>\r\n" +
                   "    <c:plotVisOnly val=\"1\"/>\r\n";
        }

        // --- Per-type inner element builders ---

        private static string BuildBarInner(string barDir, string grouping)
        {
            return "      <c:barChart>" +
                   "<c:barDir val=\"" + barDir + "\"/>" +
                   "<c:grouping val=\"" + grouping + "\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal("__DATA__") +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/>" +
                   "</c:barChart>\r\n";
        }

        private static string BuildLineInner(string grouping)
        {
            return "      <c:lineChart>" +
                   "<c:grouping val=\"" + grouping + "\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal("__DATA__") +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/>" +
                   "</c:lineChart>\r\n";
        }

        private static string BuildAreaInner(string grouping)
        {
            return "      <c:areaChart>" +
                   "<c:grouping val=\"" + grouping + "\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal("__DATA__") +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/>" +
                   "</c:areaChart>\r\n";
        }

        private static string BuildBar3DInner(string barDir, string grouping)
        {
            return "      <c:bar3DChart>" +
                   "<c:barDir val=\"" + barDir + "\"/>" +
                   "<c:grouping val=\"" + grouping + "\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal("__DATA__") +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/><c:axId val=\"3\"/>" +
                   "</c:bar3DChart>\r\n";
        }

        private static string BuildLine3DInner()
        {
            return "      <c:line3DChart>" +
                   "<c:grouping val=\"standard\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal("__DATA__") +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/><c:axId val=\"3\"/>" +
                   "</c:line3DChart>\r\n";
        }

        private static string BuildArea3DInner(string grouping)
        {
            return "      <c:area3DChart>" +
                   "<c:grouping val=\"" + grouping + "\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal("__DATA__") +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/><c:axId val=\"3\"/>" +
                   "</c:area3DChart>\r\n";
        }

        // --- Full document builders ---

        private static string BuildWithAxes(string dataRange, string chartTitle, string chartInner)
        {
            var inner = chartInner.Replace("__DATA__", dataRange);
            return Header() +
                   "  <c:chart>\r\n" +
                   TitleXml(chartTitle) +
                   "    <c:plotArea>\r\n" +
                   "      <c:layout/>\r\n" +
                   inner +
                   CatAxValAx() +
                   "    </c:plotArea>\r\n" +
                   Legend() +
                   "  </c:chart>\r\n" +
                   Footer();
        }

        private static string BuildWithAxes3D(string dataRange, string chartTitle, string chartInner)
        {
            var inner = chartInner.Replace("__DATA__", dataRange);
            return Header() +
                   "  <c:chart>\r\n" +
                   TitleXml(chartTitle) +
                   "    <c:view3D><c:rotX val=\"15\"/><c:rotY val=\"20\"/><c:perspective val=\"30\"/></c:view3D>\r\n" +
                   "    <c:plotArea>\r\n" +
                   "      <c:layout/>\r\n" +
                   inner +
                   CatAxValAx() +
                   SerAx(deleted: true) +
                   "    </c:plotArea>\r\n" +
                   Legend() +
                   "  </c:chart>\r\n" +
                   Footer();
        }

        private static string BuildPieStyle(string dataRange, string chartTitle, string elementName, int holeSize = 0)
        {
            var holeSizeAttr = holeSize > 0
                ? " <c:holeSize val=\"" + holeSize.ToString(CultureInfo.InvariantCulture) + "\"/>"
                : string.Empty;
            return Header() +
                   "  <c:chart>\r\n" +
                   TitleXml(chartTitle) +
                   "    <c:plotArea>\r\n" +
                   "      <c:layout/>\r\n" +
                   "      <c:" + elementName + ">" +
                   "<c:varyColors val=\"1\"/>\r\n" +
                   "        <c:ser>" +
                   "<c:idx val=\"0\"/><c:order val=\"0\"/>" +
                   "<c:val><c:numRef><c:f>" + dataRange + "</c:f>" +
                   "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
                   "</c:numRef></c:val>" +
                   "</c:ser>\r\n" +
                   holeSizeAttr +
                   "      </c:" + elementName + ">\r\n" +
                   "    </c:plotArea>\r\n" +
                   Legend() +
                   "  </c:chart>\r\n" +
                   Footer();
        }

        private static string BuildScatterStyle(string dataRange, string chartTitle)
        {
            return Header() +
                   "  <c:chart>\r\n" +
                   TitleXml(chartTitle) +
                   "    <c:plotArea>\r\n" +
                   "      <c:layout/>\r\n" +
                   "      <c:scatterChart>" +
                   "<c:scatterStyle val=\"marker\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   "        <c:ser>" +
                   "<c:idx val=\"0\"/><c:order val=\"0\"/>" +
                   "<c:yVal><c:numRef><c:f>" + dataRange + "</c:f>" +
                   "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
                   "</c:numRef></c:yVal>" +
                   "</c:ser>\r\n" +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/>" +
                   "</c:scatterChart>\r\n" +
                   ValAxValAx() +
                   "    </c:plotArea>\r\n" +
                   Legend() +
                   "  </c:chart>\r\n" +
                   Footer();
        }

        private static string BuildBubbleStyle(string dataRange, string chartTitle)
        {
            return Header() +
                   "  <c:chart>\r\n" +
                   TitleXml(chartTitle) +
                   "    <c:plotArea>\r\n" +
                   "      <c:layout/>\r\n" +
                   "      <c:bubbleChart>" +
                   "<c:varyColors val=\"0\"/>" +
                   "<c:showNegBubbles val=\"0\"/>\r\n" +
                   "        <c:ser>" +
                   "<c:idx val=\"0\"/><c:order val=\"0\"/>" +
                   "<c:yVal><c:numRef><c:f>" + dataRange + "</c:f>" +
                   "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
                   "</c:numRef></c:yVal>" +
                   "<c:bubbleSize><c:numRef><c:f>" + dataRange + "</c:f>" +
                   "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>" +
                   "</c:numRef></c:bubbleSize>" +
                   "</c:ser>\r\n" +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/>" +
                   "</c:bubbleChart>\r\n" +
                   ValAxValAx() +
                   "    </c:plotArea>\r\n" +
                   Legend() +
                   "  </c:chart>\r\n" +
                   Footer();
        }

        private static string BuildRadarStyle(string dataRange, string chartTitle)
        {
            return Header() +
                   "  <c:chart>\r\n" +
                   TitleXml(chartTitle) +
                   "    <c:plotArea>\r\n" +
                   "      <c:layout/>\r\n" +
                   "      <c:radarChart>" +
                   "<c:radarStyle val=\"marker\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal(dataRange) +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/>" +
                   "</c:radarChart>\r\n" +
                   CatAxValAx() +
                   "    </c:plotArea>\r\n" +
                   Legend() +
                   "  </c:chart>\r\n" +
                   Footer();
        }

        private static string BuildStockStyle(string dataRange, string chartTitle)
        {
            // Stock chart requires 4 series (Open, High, Low, Close)
            var sb = new StringBuilder();
            sb.Append(Header());
            sb.Append("  <c:chart>\r\n");
            sb.Append(TitleXml(chartTitle));
            sb.Append("    <c:plotArea>\r\n");
            sb.Append("      <c:layout/>\r\n");
            sb.Append("      <c:stockChart>\r\n");
            for (var i = 0; i < 4; i++)
            {
                sb.Append(SeriesWithVal(dataRange, i));
            }

            sb.Append("        <c:axId val=\"1\"/><c:axId val=\"2\"/>");
            sb.Append("</c:stockChart>\r\n");
            sb.Append(CatAxValAx());
            sb.Append("    </c:plotArea>\r\n");
            sb.Append(Legend());
            sb.Append("  </c:chart>\r\n");
            sb.Append(Footer());
            return sb.ToString();
        }

        private static string BuildSurface3DStyle(string dataRange, string chartTitle, bool wireframe)
        {
            var wireframeAttr = wireframe ? "<c:wireframe val=\"1\"/>" : "<c:wireframe val=\"0\"/>";
            return Header() +
                   "  <c:chart>\r\n" +
                   TitleXml(chartTitle) +
                   "    <c:view3D><c:rotX val=\"15\"/><c:rotY val=\"20\"/><c:perspective val=\"30\"/></c:view3D>\r\n" +
                   "    <c:plotArea>\r\n" +
                   "      <c:layout/>\r\n" +
                   "      <c:surface3DChart>" +
                   wireframeAttr +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal(dataRange) +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/><c:axId val=\"3\"/>" +
                   "</c:surface3DChart>\r\n" +
                   CatAxValAx() +
                   SerAx(deleted: false) +
                   "    </c:plotArea>\r\n" +
                   Legend() +
                   "  </c:chart>\r\n" +
                   Footer();
        }

        private static string BuildContourStyle(string dataRange, string chartTitle)
        {
            return Header() +
                   "  <c:chart>\r\n" +
                   TitleXml(chartTitle) +
                   "    <c:plotArea>\r\n" +
                   "      <c:layout/>\r\n" +
                   "      <c:surfaceChart>" +
                   "<c:wireframe val=\"0\"/>" +
                   "<c:varyColors val=\"0\"/>\r\n" +
                   SeriesWithVal(dataRange) +
                   "        <c:axId val=\"1\"/><c:axId val=\"2\"/><c:axId val=\"3\"/>" +
                   "</c:surfaceChart>\r\n" +
                   CatAxValAx() +
                   SerAx(deleted: false) +
                   "    </c:plotArea>\r\n" +
                   Legend() +
                   "  </c:chart>\r\n" +
                   Footer();
        }
    }
}
