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
        /// Performs style value.default.clone.
        /// </summary>
        /// <returns>The style value default cell style { get; set; } =.</returns>
        public StyleValue DefaultCellStyle { get; set; } = StyleValue.Default.Clone();
        /// <summary>
        /// Performs style value.default.clone.
        /// </summary>
        /// <returns>The list<style value> cell formats { get; } = new list<style value> {.</returns>
        public List<StyleValue> CellFormats { get; } = new List<StyleValue> { StyleValue.Default.Clone() };
        /// <summary>
        /// Performs list<style value>.
        /// </summary>
        /// <returns>The list<style value> differential formats { get; } = new.</returns>
        public List<StyleValue> DifferentialFormats { get; } = new List<StyleValue>();
        /// <summary>
        /// Performs hash set<int>.
        /// </summary>
        /// <returns>The set.</returns>
        public HashSet<int> DateStyleIndexes { get; } = new HashSet<int>();
    }
}
