using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents the internal model for an image (picture) anchored to a worksheet.
    /// Row and column values are zero-based. EMU offset values are in English Metric Units.
    /// </summary>
    public sealed class PictureModel
    {
        /// <summary>
        /// Initializes a new instance with default field values.
        /// </summary>
        public PictureModel()
        {
            Name = string.Empty;
            ImageExtension = "jpeg";
            ImageData = new byte[0];
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
        /// Gets or sets the image width extent in EMU used in spPr. Zero means compute from anchor span.
        /// </summary>
        public long ExtentCx { get; set; }

        /// <summary>
        /// Gets or sets the image height extent in EMU used in spPr. Zero means compute from anchor span.
        /// </summary>
        public long ExtentCy { get; set; }

        /// <summary>
        /// Gets or sets the lowercase file extension without dot (e.g. "jpeg", "png", "gif", "bmp").
        /// </summary>
        public string ImageExtension { get; set; }

        /// <summary>
        /// Gets or sets the raw binary bytes of the image file.
        /// </summary>
        public byte[] ImageData { get; set; }

        /// <summary>
        /// Gets or sets the original relationship ID this picture had in the drawing rels file (e.g. "rId2").
        /// Used to remap rId references in group-shape raw XML when the picture is renumbered on save.
        /// </summary>
        internal string OriginalRId { get; set; }
    }
}
