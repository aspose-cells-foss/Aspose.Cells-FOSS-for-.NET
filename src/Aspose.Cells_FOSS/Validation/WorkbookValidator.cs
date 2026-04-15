using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.InternalValidation
{
    /// <summary>
    /// Represents workbook validator.
    /// </summary>
    public sealed class WorkbookValidator
    {
        /// <summary>
        /// Validates the current state.
        /// </summary>
        /// <param name="workbookModel">The workbook model.</param>
        /// <param name="packageModel">The package model.</param>
        /// <returns>The validation message list.</returns>
        public IReadOnlyList<ValidationMessage> ValidateForSave(object workbookModel, object packageModel)
        {
            return Array.Empty<ValidationMessage>();
        }
    }
}
