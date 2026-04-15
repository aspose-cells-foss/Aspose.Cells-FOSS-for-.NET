using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents workbook properties model.
    /// </summary>
    public sealed class WorkbookPropertiesModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookPropertiesModel"/> class.
        /// </summary>
        public WorkbookPropertiesModel()
        {
            Protection = new WorkbookProtectionModel();
            View = new WorkbookViewModel();
            Calculation = new CalculationPropertiesModel();
        }

        /// <summary>
        /// Gets or sets the code name.
        /// </summary>
        public string CodeName { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the show objects.
        /// </summary>
        public string ShowObjects { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether filter privacy.
        /// </summary>
        public bool FilterPrivacy { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether show border unselected tables.
        /// </summary>
        public bool ShowBorderUnselectedTables { get; set; } = true;
        /// <summary>
        /// Gets or sets a value indicating whether show ink annotation.
        /// </summary>
        public bool ShowInkAnnotation { get; set; } = true;
        /// <summary>
        /// Gets or sets a value indicating whether backup file.
        /// </summary>
        public bool BackupFile { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether save external link values.
        /// </summary>
        public bool SaveExternalLinkValues { get; set; } = true;
        /// <summary>
        /// Gets or sets the update links.
        /// </summary>
        public string UpdateLinks { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether hide pivot field list.
        /// </summary>
        public bool HidePivotFieldList { get; set; }
        /// <summary>
        /// Gets or sets the default theme version.
        /// </summary>
        public int? DefaultThemeVersion { get; set; }
        /// <summary>
        /// Gets the protection.
        /// </summary>
        public WorkbookProtectionModel Protection { get; }
        /// <summary>
        /// Gets the view.
        /// </summary>
        public WorkbookViewModel View { get; }
        /// <summary>
        /// Gets the calculation.
        /// </summary>
        public CalculationPropertiesModel Calculation { get; }

        /// <summary>
        /// Copies values from the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        public void CopyFrom(WorkbookPropertiesModel source)
        {
            CodeName = source.CodeName;
            ShowObjects = source.ShowObjects;
            FilterPrivacy = source.FilterPrivacy;
            ShowBorderUnselectedTables = source.ShowBorderUnselectedTables;
            ShowInkAnnotation = source.ShowInkAnnotation;
            BackupFile = source.BackupFile;
            SaveExternalLinkValues = source.SaveExternalLinkValues;
            UpdateLinks = source.UpdateLinks;
            HidePivotFieldList = source.HidePivotFieldList;
            DefaultThemeVersion = source.DefaultThemeVersion;
            Protection.CopyFrom(source.Protection);
            View.CopyFrom(source.View);
            Calculation.CopyFrom(source.Calculation);
        }

        /// <summary>
        /// Performs has workbook properties state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasWorkbookPropertiesState()
        {
            return !string.IsNullOrEmpty(CodeName)
                || !string.IsNullOrEmpty(ShowObjects)
                || FilterPrivacy
                || !ShowBorderUnselectedTables
                || !ShowInkAnnotation
                || BackupFile
                || !SaveExternalLinkValues
                || !string.IsNullOrEmpty(UpdateLinks)
                || HidePivotFieldList
                || DefaultThemeVersion.HasValue;
        }
    }
}
