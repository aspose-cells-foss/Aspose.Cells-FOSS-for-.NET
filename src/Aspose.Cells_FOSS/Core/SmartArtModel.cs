namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents a SmartArt diagram anchored to a worksheet. Row/column values are zero-based and
    /// EMU offsets are in English Metric Units. <see cref="RawDrawingXml"/> holds the diagram's
    /// pre-laid-out drawing part (dsp:drawing), whose shapes carry absolute positions relative to the
    /// diagram frame - so rendering does not require a SmartArt layout engine.
    /// </summary>
    internal sealed class SmartArtModel
    {
        public string Name { get; set; } = string.Empty;

        public int UpperLeftRow { get; set; }
        public int UpperLeftColumn { get; set; }
        public long UpperLeftColumnOffset { get; set; }
        public long UpperLeftRowOffset { get; set; }
        public int LowerRightRow { get; set; }
        public int LowerRightColumn { get; set; }
        public long LowerRightColumnOffset { get; set; }
        public long LowerRightRowOffset { get; set; }

        /// <summary>Raw XML of the diagram drawing part (dsp:drawing) with the laid-out shapes.</summary>
        public string RawDrawingXml { get; set; }
    }
}
