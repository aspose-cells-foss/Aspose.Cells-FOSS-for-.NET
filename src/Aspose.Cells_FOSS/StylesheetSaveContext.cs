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
    internal sealed class StylesheetSaveContext
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StylesheetSaveContext"/> class.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="styleIndices">The style indices.</param>
        /// <param name="differentialStyleIndices">The differential style indices.</param>
        /// <param name="differentialFormatCount">The differential format count.</param>
        /// <param name="hasStyles">The has styles.</param>
        public StylesheetSaveContext(XDocument document, IReadOnlyDictionary<string, int> styleIndices, IReadOnlyDictionary<string, int> differentialStyleIndices, int differentialFormatCount, bool hasStyles)
        {
            Document = document;
            _styleIndices = styleIndices;
            _differentialStyleIndices = differentialStyleIndices;
            DifferentialFormatCount = differentialFormatCount;
            HasStyles = hasStyles;
        }

        private readonly IReadOnlyDictionary<string, int> _styleIndices;
        private readonly IReadOnlyDictionary<string, int> _differentialStyleIndices;
        /// <summary>
        /// Gets the document.
        /// </summary>
        public XDocument Document { get; }
        /// <summary>
        /// Gets the differential format count.
        /// </summary>
        public int DifferentialFormatCount { get; }
        /// <summary>
        /// Gets a value indicating whether styles.
        /// </summary>
        public bool HasStyles { get; }

        /// <summary>
        /// Gets the style index.
        /// </summary>
        /// <param name="record">The record.</param>
        /// <returns>The int.</returns>
        public int GetStyleIndex(CellRecord record)
        {
            var style = XlsxWorkbookStyles.GetStyleForSerialization(record);
            int index;
            return _styleIndices.TryGetValue(XlsxWorkbookStyles.GetStyleKey(style), out index) ? index : 0;
        }

        /// <summary>
        /// Gets the differential style index.
        /// </summary>
        /// <param name="condition">The condition.</param>
        /// <returns>The int.</returns>
        public int? GetDifferentialStyleIndex(FormatConditionModel condition)
        {
            if (XlsxWorkbookStyles.StylesEqual(condition.Style, StyleValue.Default))
            {
                return null;
            }

            int differentialIndex;
            if (_differentialStyleIndices.TryGetValue(XlsxWorkbookStyles.GetStyleKey(condition.Style), out differentialIndex))
            {
                return differentialIndex;
            }
            return null;
        }
    }
}
