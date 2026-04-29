using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;
using static Aspose.Cells_FOSS.XlsxWorkbookStylesXml;

namespace Aspose.Cells_FOSS
{
    internal sealed class CellFormatValue
    {
        /// <summary>
        /// Gets or sets the num fmt id.
        /// </summary>
        public int NumFmtId { get; set; }
        /// <summary>
        /// Gets or sets the font id.
        /// </summary>
        public int FontId { get; set; }
        /// <summary>
        /// Gets or sets the fill id.
        /// </summary>
        public int FillId { get; set; }
        /// <summary>
        /// Gets or sets the border id.
        /// </summary>
        public int BorderId { get; set; }
        /// <summary>
        /// Performs alignment value.
        /// </summary>
        /// <returns>The alignment value alignment { get; set; } = new.</returns>
        public AlignmentValue Alignment { get; set; } = new AlignmentValue();
        /// <summary>
        /// Performs protection value.
        /// </summary>
        /// <returns>The protection value protection { get; set; } = new.</returns>
        public ProtectionValue Protection { get; set; } = new ProtectionValue();
    }
}
