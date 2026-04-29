using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Xml
{
    /// <summary>
    /// Represents workbook xml mapper.
    /// </summary>
    public sealed class WorkbookXmlMapper
    {
        /// <summary>
        /// Reads data from the specified source.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="workbookModel">The workbook model.</param>
        /// <param name="packageModel">The package model.</param>
        public void Read(Stream stream, object workbookModel, object packageModel)
        {
            throw new NotSupportedException("SpreadsheetML reading is not implemented in this initial solution skeleton.");
        }

        /// <summary>
        /// Writes data to the specified target.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="workbookModel">The workbook model.</param>
        /// <param name="packageModel">The package model.</param>
        public void Write(Stream stream, object workbookModel, object packageModel)
        {
            throw new NotSupportedException("SpreadsheetML writing is not implemented in this initial solution skeleton.");
        }
    }
}
