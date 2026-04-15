using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents workbook properties.
    /// </summary>
    public sealed class WorkbookProperties
    {
        private readonly WorkbookModel _workbookModel;
        private readonly WorkbookPropertiesModel _model;
        private readonly WorkbookProtection _protection;
        private readonly WorkbookView _view;
        private readonly CalculationProperties _calculation;

        internal WorkbookProperties(WorkbookModel workbookModel)
        {
            _workbookModel = workbookModel;
            _model = workbookModel.Properties;
            _protection = new WorkbookProtection(_model.Protection);
            _view = new WorkbookView(workbookModel);
            _calculation = new CalculationProperties(_model.Calculation);
        }

        /// <summary>
        /// Gets or sets the code name.
        /// </summary>
        public string CodeName
        {
            get
            {
                return _model.CodeName;
            }
            set
            {
                _model.CodeName = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the show objects.
        /// </summary>
        public string ShowObjects
        {
            get
            {
                return string.IsNullOrEmpty(_model.ShowObjects) ? "all" : _model.ShowObjects;
            }
            set
            {
                _model.ShowObjects = WorkbookPropertySupport.NormalizeShowObjects(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether filter privacy.
        /// </summary>
        public bool FilterPrivacy
        {
            get
            {
                return _model.FilterPrivacy;
            }
            set
            {
                _model.FilterPrivacy = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether show border unselected tables.
        /// </summary>
        public bool ShowBorderUnselectedTables
        {
            get
            {
                return _model.ShowBorderUnselectedTables;
            }
            set
            {
                _model.ShowBorderUnselectedTables = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether show ink annotation.
        /// </summary>
        public bool ShowInkAnnotation
        {
            get
            {
                return _model.ShowInkAnnotation;
            }
            set
            {
                _model.ShowInkAnnotation = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether backup file.
        /// </summary>
        public bool BackupFile
        {
            get
            {
                return _model.BackupFile;
            }
            set
            {
                _model.BackupFile = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether save external link values.
        /// </summary>
        public bool SaveExternalLinkValues
        {
            get
            {
                return _model.SaveExternalLinkValues;
            }
            set
            {
                _model.SaveExternalLinkValues = value;
            }
        }

        /// <summary>
        /// Gets or sets the update links.
        /// </summary>
        public string UpdateLinks
        {
            get
            {
                return string.IsNullOrEmpty(_model.UpdateLinks) ? "userSet" : _model.UpdateLinks;
            }
            set
            {
                _model.UpdateLinks = WorkbookPropertySupport.NormalizeUpdateLinks(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether hide pivot field list.
        /// </summary>
        public bool HidePivotFieldList
        {
            get
            {
                return _model.HidePivotFieldList;
            }
            set
            {
                _model.HidePivotFieldList = value;
            }
        }

        /// <summary>
        /// Gets or sets the default theme version.
        /// </summary>
        public int? DefaultThemeVersion
        {
            get
            {
                return _model.DefaultThemeVersion;
            }
            set
            {
                if (value.HasValue && value.Value < 0)
                {
                    throw new CellsException("DefaultThemeVersion must be non-negative.");
                }

                _model.DefaultThemeVersion = value;
            }
        }

        /// <summary>
        /// Gets the protection.
        /// </summary>
        public WorkbookProtection Protection
        {
            get
            {
                return _protection;
            }
        }

        /// <summary>
        /// Gets the view.
        /// </summary>
        public WorkbookView View
        {
            get
            {
                return _view;
            }
        }

        /// <summary>
        /// Gets the calculation.
        /// </summary>
        public CalculationProperties Calculation
        {
            get
            {
                return _calculation;
            }
        }
    }
}
