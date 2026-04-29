using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents the ordered collection of columns in an Excel table.
    /// </summary>
    public sealed class ListColumnCollection
    {
        private readonly ListObjectModel _model;

        internal ListColumnCollection(ListObjectModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets the number of columns in the table.
        /// </summary>
        public int Count
        {
            get
            {
                return _model.Columns.Count;
            }
        }

        /// <summary>
        /// Gets the column at the specified zero-based index.
        /// </summary>
        public ListColumn this[int index]
        {
            get
            {
                if (index < 0 || index >= _model.Columns.Count)
                {
                    throw new CellsException("Column index " + index + " is out of range.");
                }

                return new ListColumn(_model.Columns[index]);
            }
        }
    }
}
