namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents the internal model for a drawing object (shape) anchored to a worksheet.
    /// Row and column values are zero-based. EMU offset values are in English Metric Units.
    /// </summary>
    public sealed class ShapeModel
    {
        /// <summary>
        /// Initializes a new instance with default field values.
        /// </summary>
        public ShapeModel()
        {
            Name = string.Empty;
            GeometryType = "rect";
        }

        /// <summary>
        /// Gets or sets the display name shown in Excel's name box.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the zero-based row index of the upper-left anchor cell.
        /// </summary>
        public int UpperLeftRow { get; set; }

        /// <summary>
        /// Gets or sets the zero-based column index of the upper-left anchor cell.
        /// </summary>
        public int UpperLeftColumn { get; set; }

        /// <summary>
        /// Gets or sets the EMU offset from the left edge of the upper-left anchor cell.
        /// </summary>
        public long UpperLeftColumnOffset { get; set; }

        /// <summary>
        /// Gets or sets the EMU offset from the top edge of the upper-left anchor cell.
        /// </summary>
        public long UpperLeftRowOffset { get; set; }

        /// <summary>
        /// Gets or sets the zero-based row index of the lower-right anchor cell.
        /// </summary>
        public int LowerRightRow { get; set; }

        /// <summary>
        /// Gets or sets the zero-based column index of the lower-right anchor cell.
        /// </summary>
        public int LowerRightColumn { get; set; }

        /// <summary>
        /// Gets or sets the EMU offset from the left edge of the lower-right anchor cell.
        /// </summary>
        public long LowerRightColumnOffset { get; set; }

        /// <summary>
        /// Gets or sets the EMU offset from the top edge of the lower-right anchor cell.
        /// </summary>
        public long LowerRightRowOffset { get; set; }

        /// <summary>
        /// Gets or sets the shape width extent in EMU used in spPr. Zero means compute from anchor span.
        /// </summary>
        public long ExtentCx { get; set; }

        /// <summary>
        /// Gets or sets the shape height extent in EMU used in spPr. Zero means compute from anchor span.
        /// </summary>
        public long ExtentCy { get; set; }

        /// <summary>
        /// Gets or sets the DrawingML preset geometry type (e.g. "rect", "rightArrow", "star12").
        /// </summary>
        public string GeometryType { get; set; }

        /// <summary>
        /// Gets or sets the preserved outer XML of the &lt;xdr:style&gt; element, or null if absent.
        /// </summary>
        public string RawStyleXml { get; set; }

        /// <summary>
        /// Gets or sets the preserved outer XML of the &lt;xdr:txBody&gt; element, or null if absent.
        /// </summary>
        public string RawTxBodyXml { get; set; }

        /// <summary>
        /// Gets or sets the preserved outer XML of the shape element (e.g. &lt;xdr:cxnSp&gt;) for
        /// shapes that cannot be represented in the standard ShapeModel fields. When set, this XML is
        /// emitted verbatim as the shape element inside the anchor. Null for regular &lt;xdr:sp&gt; shapes.
        /// </summary>
        public string RawElementXml { get; set; }
    }
}
