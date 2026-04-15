using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal sealed class WorksheetDefinedNamesState
    {
        /// <summary>
        /// Gets or sets the print area.
        /// </summary>
        public string PrintArea { get; set; }
        /// <summary>
        /// Gets or sets the print title rows.
        /// </summary>
        public string PrintTitleRows { get; set; }
        /// <summary>
        /// Gets or sets the print title columns.
        /// </summary>
        public string PrintTitleColumns { get; set; }
    }
}
