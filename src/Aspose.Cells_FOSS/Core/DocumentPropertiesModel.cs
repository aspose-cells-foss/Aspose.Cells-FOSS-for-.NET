using System.IO;
using System.Collections.Generic;
using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents document properties model.
    /// </summary>
    public sealed class DocumentPropertiesModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentPropertiesModel"/> class.
        /// </summary>
        public DocumentPropertiesModel()
        {
            Core = new CoreDocumentPropertiesModel();
            Extended = new ExtendedDocumentPropertiesModel();
        }

        /// <summary>
        /// Gets the core.
        /// </summary>
        public CoreDocumentPropertiesModel Core { get; }
        /// <summary>
        /// Gets the extended.
        /// </summary>
        public ExtendedDocumentPropertiesModel Extended { get; }

        /// <summary>
        /// Copies values from the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        public void CopyFrom(DocumentPropertiesModel source)
        {
            Core.CopyFrom(source.Core);
            Extended.CopyFrom(source.Extended);
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return Core.HasStoredState() || Extended.HasStoredState();
        }
    }
}
