using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents core document properties.
    /// </summary>
    public sealed class CoreDocumentProperties
    {
        private readonly CoreDocumentPropertiesModel _model;

        internal CoreDocumentProperties(CoreDocumentPropertiesModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets the title.
        /// </summary>
        public string Title
        {
            get
            {
                return _model.Title;
            }
            set
            {
                _model.Title = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the subject.
        /// </summary>
        public string Subject
        {
            get
            {
                return _model.Subject;
            }
            set
            {
                _model.Subject = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the creator.
        /// </summary>
        public string Creator
        {
            get
            {
                return _model.Creator;
            }
            set
            {
                _model.Creator = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the keywords.
        /// </summary>
        public string Keywords
        {
            get
            {
                return _model.Keywords;
            }
            set
            {
                _model.Keywords = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the description.
        /// </summary>
        public string Description
        {
            get
            {
                return _model.Description;
            }
            set
            {
                _model.Description = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the last modified by.
        /// </summary>
        public string LastModifiedBy
        {
            get
            {
                return _model.LastModifiedBy;
            }
            set
            {
                _model.LastModifiedBy = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the revision.
        /// </summary>
        public string Revision
        {
            get
            {
                return _model.Revision;
            }
            set
            {
                _model.Revision = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the category.
        /// </summary>
        public string Category
        {
            get
            {
                return _model.Category;
            }
            set
            {
                _model.Category = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the content status.
        /// </summary>
        public string ContentStatus
        {
            get
            {
                return _model.ContentStatus;
            }
            set
            {
                _model.ContentStatus = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the created.
        /// </summary>
        public DateTime? Created
        {
            get
            {
                return _model.Created;
            }
            set
            {
                _model.Created = value;
            }
        }

        /// <summary>
        /// Gets or sets the modified.
        /// </summary>
        public DateTime? Modified
        {
            get
            {
                return _model.Modified;
            }
            set
            {
                _model.Modified = value;
            }
        }
    }
}
