using System.Linq;
using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents style repository.
    /// </summary>
    public sealed class StyleRepository
    {
        /// <summary>
        /// Normalizes the specified value.
        /// </summary>
        /// <param name="style">The style.</param>
        /// <returns>The style value.</returns>
        public StyleValue Normalize(StyleValue style)
        {
            return style.Clone();
        }
    }
}
