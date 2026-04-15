using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents core document properties model.
    /// </summary>
    public sealed class CoreDocumentPropertiesModel
    {
        /// <summary>
        /// Gets or sets the title.
        /// </summary>
        public string Title { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the subject.
        /// </summary>
        public string Subject { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the creator.
        /// </summary>
        public string Creator { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the keywords.
        /// </summary>
        public string Keywords { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the description.
        /// </summary>
        public string Description { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the last modified by.
        /// </summary>
        public string LastModifiedBy { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the revision.
        /// </summary>
        public string Revision { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the category.
        /// </summary>
        public string Category { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the content status.
        /// </summary>
        public string ContentStatus { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the created.
        /// </summary>
        public DateTime? Created { get; set; }
        /// <summary>
        /// Gets or sets the modified.
        /// </summary>
        public DateTime? Modified { get; set; }

        /// <summary>
        /// Copies values from the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        public void CopyFrom(CoreDocumentPropertiesModel source)
        {
            Title = source.Title;
            Subject = source.Subject;
            Creator = source.Creator;
            Keywords = source.Keywords;
            Description = source.Description;
            LastModifiedBy = source.LastModifiedBy;
            Revision = source.Revision;
            Category = source.Category;
            ContentStatus = source.ContentStatus;
            Created = source.Created;
            Modified = source.Modified;
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return !string.IsNullOrEmpty(Title)
                || !string.IsNullOrEmpty(Subject)
                || !string.IsNullOrEmpty(Creator)
                || !string.IsNullOrEmpty(Keywords)
                || !string.IsNullOrEmpty(Description)
                || !string.IsNullOrEmpty(LastModifiedBy)
                || !string.IsNullOrEmpty(Revision)
                || !string.IsNullOrEmpty(Category)
                || !string.IsNullOrEmpty(ContentStatus)
                || Created.HasValue
                || Modified.HasValue;
        }
    }
}
