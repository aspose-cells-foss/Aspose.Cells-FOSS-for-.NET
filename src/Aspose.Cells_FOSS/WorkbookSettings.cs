using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents workbook-level settings that affect date handling and display formatting.
    /// </summary>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// workbook.Settings.Date1904 = true;
    /// workbook.Settings.Culture = CultureInfo.GetCultureInfo("fr-FR");
    /// </code>
    /// </example>
    public sealed class WorkbookSettings
    {
        private readonly WorkbookSettingsModel _model;

        internal WorkbookSettings(WorkbookSettingsModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets whether the workbook uses the 1904 date system.
        /// </summary>
        public bool Date1904
        {
            get
            {
                return _model.DateSystem == Aspose.Cells_FOSS.Core.DateSystem.Mac1904;
            }
            set
            {
                _model.DateSystem = value ? Aspose.Cells_FOSS.Core.DateSystem.Mac1904 : Aspose.Cells_FOSS.Core.DateSystem.Windows1900;
            }
        }

        /// <summary>
        /// Gets or sets the culture used for display-string formatting.
        /// </summary>
        public CultureInfo Culture
        {
            get
            {
                // Expose a defensive copy so callers can mutate the returned CultureInfo
                // without changing workbook formatting rules until they assign it back.
                return (CultureInfo)_model.DisplayCulture.Clone();
            }
            set
            {
                if (value == null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                _model.DisplayCulture = (CultureInfo)value.Clone();
            }
        }
    }
}
