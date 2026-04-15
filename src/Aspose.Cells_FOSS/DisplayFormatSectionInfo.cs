using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Aspose.Cells_FOSS
{
    internal sealed class DisplayFormatSectionInfo
    {
        /// <summary>
        /// Gets or sets the raw.
        /// </summary>
        public string Raw { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether condition.
        /// </summary>
        public bool HasCondition { get; set; }
        /// <summary>
        /// Gets or sets the condition operator.
        /// </summary>
        public string ConditionOperator { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets the condition value.
        /// </summary>
        public double ConditionValue { get; set; }
    }
}
