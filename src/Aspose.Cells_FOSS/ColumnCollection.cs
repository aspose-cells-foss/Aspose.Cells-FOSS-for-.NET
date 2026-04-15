using System.IO;
using System.Collections.Generic;
using System;
using System.Linq;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a collection of column objects.
    /// </summary>
    public sealed class ColumnCollection
    {
        private readonly Worksheet _worksheet;

        internal ColumnCollection(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        /// <summary>
        /// Gets the element at the specified zero-based index.
        /// </summary>
        public Column this[int index]
        {
            get
            {
                if (index < 0) throw new CellsException("Column index must be non-negative.");
                return new Column(_worksheet, index);
            }
        }
    }
}
