using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents a chart embedded in a worksheet drawing layer.
    /// </summary>
    public sealed class ChartModel
    {
        /// <summary>Gets or sets the chart display name.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Gets or sets the detected chart type.</summary>
        public ChartType ChartType { get; set; } = ChartType.Unknown;

        /// <summary>Gets or sets the zero-based upper-left row of the chart anchor.</summary>
        public int UpperLeftRow { get; set; }

        /// <summary>Gets or sets the zero-based upper-left column of the chart anchor.</summary>
        public int UpperLeftColumn { get; set; }

        /// <summary>Gets or sets the upper-left row offset in EMU.</summary>
        public long UpperLeftRowOffset { get; set; }

        /// <summary>Gets or sets the upper-left column offset in EMU.</summary>
        public long UpperLeftColumnOffset { get; set; }

        /// <summary>Gets or sets the zero-based lower-right row of the chart anchor.</summary>
        public int LowerRightRow { get; set; }

        /// <summary>Gets or sets the zero-based lower-right column of the chart anchor.</summary>
        public int LowerRightColumn { get; set; }

        /// <summary>Gets or sets the lower-right row offset in EMU.</summary>
        public long LowerRightRowOffset { get; set; }

        /// <summary>Gets or sets the lower-right column offset in EMU.</summary>
        public long LowerRightColumnOffset { get; set; }

        /// <summary>Gets or sets the width extent in EMU (0 = compute from anchor).</summary>
        public long ExtentCx { get; set; }

        /// <summary>Gets or sets the height extent in EMU (0 = compute from anchor).</summary>
        public long ExtentCy { get; set; }

        /// <summary>Gets or sets the raw XML content of the chart definition file (xl/charts/chart{N}.xml or chartEx{N}.xml).</summary>
        public string RawChartXml { get; set; }

        /// <summary>Gets or sets the raw XML of the graphicFrame container element (mc:AlternateContent for chartex charts). Used verbatim on save with the rId substituted.</summary>
        public string RawGraphicFrameXml { get; set; }

        /// <summary>Gets or sets the original relationship ID from the drawing (e.g. "rId1"), used to substitute the new rId into RawGraphicFrameXml on save.</summary>
        public string OriginalRId { get; set; }

        /// <summary>Gets or sets whether this chart uses the chartex format (xl/charts/chartEx{N}.xml, different relationship and content types).</summary>
        public bool IsChartEx { get; set; }

        /// <summary>Gets the companion files referenced by the chart (style, colors, etc.).</summary>
        public List<ChartCompanionFile> CompanionFiles { get; } = new List<ChartCompanionFile>();
    }

    /// <summary>
    /// Represents a companion file (style, colors) linked from a chart's relationship file.
    /// </summary>
    public sealed class ChartCompanionFile
    {
        /// <summary>Gets or sets the relationship ID (e.g. "rId1").</summary>
        public string RelationshipId { get; set; }

        /// <summary>Gets or sets the relationship type URI.</summary>
        public string RelationshipType { get; set; }

        /// <summary>Gets or sets the base file name (e.g. "style1.xml").</summary>
        public string FileName { get; set; }

        /// <summary>Gets or sets the raw XML content of the companion file.</summary>
        public string RawContent { get; set; }
    }
}
