using System.IO;
using System.Collections.Generic;
using System;
using System.Linq;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a collection of row objects.
    /// </summary>
    public sealed class RowCollection
    {
        private readonly Worksheet _worksheet;

        internal RowCollection(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        /// <summary>
        /// Gets the element at the specified zero-based index.
        /// </summary>
        public Row this[int index]
        {
            get
            {
                if (index < 0) throw new CellsException("Row index must be non-negative.");
                return new Row(_worksheet, index);
            }
        }
    }
}
