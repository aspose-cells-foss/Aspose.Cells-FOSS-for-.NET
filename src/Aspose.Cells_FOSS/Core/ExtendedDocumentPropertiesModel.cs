using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents extended document properties model.
    /// </summary>
    public sealed class ExtendedDocumentPropertiesModel
    {
        /// <summary>
        /// Gets or sets the application.
        /// </summary>
        public string Application { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the app version.
        /// </summary>
        public string AppVersion { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the company.
        /// </summary>
        public string Company { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the manager.
        /// </summary>
        public string Manager { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the doc security.
        /// </summary>
        public int? DocSecurity { get; set; }
        /// <summary>
        /// Gets or sets the hyperlink base.
        /// </summary>
        public string HyperlinkBase { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether scale crop.
        /// </summary>
        public bool? ScaleCrop { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether links up to date.
        /// </summary>
        public bool? LinksUpToDate { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether shared doc.
        /// </summary>
        public bool? SharedDoc { get; set; }

        /// <summary>
        /// Copies values from the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        public void CopyFrom(ExtendedDocumentPropertiesModel source)
        {
            Application = source.Application;
            AppVersion = source.AppVersion;
            Company = source.Company;
            Manager = source.Manager;
            DocSecurity = source.DocSecurity;
            HyperlinkBase = source.HyperlinkBase;
            ScaleCrop = source.ScaleCrop;
            LinksUpToDate = source.LinksUpToDate;
            SharedDoc = source.SharedDoc;
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return !string.IsNullOrEmpty(Application)
                || !string.IsNullOrEmpty(AppVersion)
                || !string.IsNullOrEmpty(Company)
                || !string.IsNullOrEmpty(Manager)
                || DocSecurity.HasValue
                || !string.IsNullOrEmpty(HyperlinkBase)
                || ScaleCrop.HasValue
                || LinksUpToDate.HasValue
                || SharedDoc.HasValue;
        }
    }
}
