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
    internal sealed class StylesheetLoadContext
    {
        /// <summary>
        /// Gets or sets the default cell style.
        /// </summary>
        public StyleValue DefaultCellStyle { get; set; } = StyleValue.Default.Clone();
        /// <summary>
        /// Gets the loaded cell format styles (xf records).
        /// </summary>
        public List<StyleValue> CellFormats { get; } = new List<StyleValue> { StyleValue.Default.Clone() };
        /// <summary>
        /// Gets the loaded differential styles (dxf records).
        /// </summary>
        public List<StyleValue> DifferentialFormats { get; } = new List<StyleValue>();
        /// <summary>
        /// Gets style indexes that should be interpreted as date/time styles.
        /// </summary>
        public HashSet<int> DateStyleIndexes { get; } = new HashSet<int>();
    }
}
