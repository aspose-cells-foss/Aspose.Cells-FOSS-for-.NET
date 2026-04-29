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
    internal sealed class FillValue
    {
        /// <summary>
        /// Gets or sets the pattern.
        /// </summary>
        public FillPatternKind Pattern { get; set; }
        /// <summary>
        /// Gets or sets the foreground color.
        /// </summary>
        public ColorValue ForegroundColor { get; set; }
        /// <summary>
        /// Gets or sets the background color.
        /// </summary>
        public ColorValue BackgroundColor { get; set; }
    }
}
