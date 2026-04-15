using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    internal sealed class ConditionalFormattingModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConditionalFormattingModel"/> class.
        /// </summary>
        public ConditionalFormattingModel()
        {
            Areas = new List<CellArea>();
            Conditions = new List<FormatConditionModel>();
        }

        /// <summary>
        /// Gets the areas.
        /// </summary>
        public List<CellArea> Areas { get; }
        /// <summary>
        /// Gets the conditions.
        /// </summary>
        public List<FormatConditionModel> Conditions { get; }
    }
}
