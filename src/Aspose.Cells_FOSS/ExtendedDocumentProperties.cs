using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents extended document properties.
    /// </summary>
    public sealed class ExtendedDocumentProperties
    {
        private readonly ExtendedDocumentPropertiesModel _model;

        internal ExtendedDocumentProperties(ExtendedDocumentPropertiesModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets the application.
        /// </summary>
        public string Application
        {
            get
            {
                return _model.Application;
            }
            set
            {
                _model.Application = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the app version.
        /// </summary>
        public string AppVersion
        {
            get
            {
                return _model.AppVersion;
            }
            set
            {
                _model.AppVersion = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the company.
        /// </summary>
        public string Company
        {
            get
            {
                return _model.Company;
            }
            set
            {
                _model.Company = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the manager.
        /// </summary>
        public string Manager
        {
            get
            {
                return _model.Manager;
            }
            set
            {
                _model.Manager = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the doc security.
        /// </summary>
        public int DocSecurity
        {
            get
            {
                return _model.DocSecurity ?? 0;
            }
            set
            {
                if (value < 0)
                {
                    throw new CellsException("DocSecurity must be non-negative.");
                }

                _model.DocSecurity = value;
            }
        }

        /// <summary>
        /// Gets or sets the hyperlink base.
        /// </summary>
        public string HyperlinkBase
        {
            get
            {
                return _model.HyperlinkBase;
            }
            set
            {
                _model.HyperlinkBase = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether scale crop.
        /// </summary>
        public bool ScaleCrop
        {
            get
            {
                return _model.ScaleCrop ?? false;
            }
            set
            {
                _model.ScaleCrop = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether links up to date.
        /// </summary>
        public bool LinksUpToDate
        {
            get
            {
                return _model.LinksUpToDate ?? false;
            }
            set
            {
                _model.LinksUpToDate = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether shared doc.
        /// </summary>
        public bool SharedDoc
        {
            get
            {
                return _model.SharedDoc ?? false;
            }
            set
            {
                _model.SharedDoc = value;
            }
        }
    }
}
