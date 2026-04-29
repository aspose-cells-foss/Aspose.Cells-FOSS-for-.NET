using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents workbook protection.
    /// </summary>
    public sealed class WorkbookProtection
    {
        private readonly WorkbookProtectionModel _model;

        internal WorkbookProtection(WorkbookProtectionModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets a value indicating whether lock structure.
        /// </summary>
        public bool LockStructure
        {
            get
            {
                return _model.LockStructure;
            }
            set
            {
                _model.LockStructure = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether lock windows.
        /// </summary>
        public bool LockWindows
        {
            get
            {
                return _model.LockWindows;
            }
            set
            {
                _model.LockWindows = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether lock revision.
        /// </summary>
        public bool LockRevision
        {
            get
            {
                return _model.LockRevision;
            }
            set
            {
                _model.LockRevision = value;
            }
        }

        /// <summary>
        /// Gets or sets the workbook password.
        /// </summary>
        public string WorkbookPassword
        {
            get
            {
                return _model.WorkbookPassword;
            }
            set
            {
                _model.WorkbookPassword = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the revisions password.
        /// </summary>
        public string RevisionsPassword
        {
            get
            {
                return _model.RevisionsPassword;
            }
            set
            {
                _model.RevisionsPassword = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets a value indicating whether protected.
        /// </summary>
        public bool IsProtected
        {
            get
            {
                return _model.HasStoredState();
            }
        }
    }
}
