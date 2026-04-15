using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents workbook settings model.
    /// </summary>
    public sealed class WorkbookSettingsModel
    {
        /// <summary>
        /// Gets or sets the date system.
        /// </summary>
        public DateSystem DateSystem { get; set; } = DateSystem.Windows1900;

        /// <summary>
        /// Gets or sets the display culture.
        /// </summary>
        public CultureInfo DisplayCulture { get; set; } = CultureInfo.InvariantCulture;
    }
}
