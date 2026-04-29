using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents the collection of Excel tables on a worksheet.
    /// </summary>
    public sealed class ListObjectCollection : IListObjectOwner
    {
        private readonly WorksheetModel _worksheetModel;
        private readonly List<ListObjectModel> _models;

        internal ListObjectCollection(WorksheetModel worksheetModel)
        {
            _worksheetModel = worksheetModel;
            _models = worksheetModel.ListObjects;
        }

        /// <summary>
        /// Gets the number of tables on the worksheet.
        /// </summary>
        public int Count
        {
            get
            {
                return _models.Count;
            }
        }

        /// <summary>
        /// Gets the table at the specified zero-based index.
        /// </summary>
        public ListObject this[int index]
        {
            get
            {
                if (index < 0 || index >= _models.Count)
                {
                    throw new CellsException("Table index " + index + " is out of range.");
                }

                return new ListObject(_models[index], _worksheetModel, this);
            }
        }

        /// <summary>
        /// Gets the table with the specified display name (case-insensitive).
        /// </summary>
        public ListObject this[string name]
        {
            get
            {
                for (var i = 0; i < _models.Count; i++)
                {
                    if (string.Equals(_models[i].DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return new ListObject(_models[i], _worksheetModel, this);
                    }
                }

                throw new CellsException("No table with display name '" + name + "' was found.");
            }
        }

        /// <summary>
        /// Adds a new table covering the specified zero-based range and returns its index.
        /// </summary>
        public int Add(int startRow, int startColumn, int endRow, int endColumn, bool hasHeaders)
        {
            ListObjectSupport.ValidateRange(startRow, startColumn, endRow, endColumn);
            ListObjectSupport.ValidateNoOverlap(_models, startRow, startColumn, endRow, endColumn, -1);
            var tableNumber = _models.Count + 1;
            var model = ListObjectSupport.CreateModel(_worksheetModel, startRow, startColumn, endRow, endColumn, hasHeaders, tableNumber);
            ListObjectSupport.ValidateUniqueDisplayName(_models, model.DisplayName, -1);
            _models.Add(model);
            return _models.Count - 1;
        }

        /// <summary>
        /// Adds a new table covering the specified A1-notation cell range and returns its index.
        /// </summary>
        public int Add(string startCellName, string endCellName, bool hasHeaders)
        {
            CellAddress startCell;
            CellAddress endCell;
            try
            {
                startCell = CellAddress.Parse(startCellName);
            }
            catch (ArgumentException)
            {
                throw new CellsException("The start cell reference '" + startCellName + "' is invalid.");
            }

            try
            {
                endCell = CellAddress.Parse(endCellName);
            }
            catch (ArgumentException)
            {
                throw new CellsException("The end cell reference '" + endCellName + "' is invalid.");
            }

            return Add(startCell.RowIndex, startCell.ColumnIndex, endCell.RowIndex, endCell.ColumnIndex, hasHeaders);
        }

        /// <summary>
        /// Removes the table at the specified zero-based index, leaving cell data in place.
        /// </summary>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _models.Count)
            {
                throw new CellsException("Table index " + index + " is out of range.");
            }

            _models.RemoveAt(index);
        }

        void IListObjectOwner.ValidateUniqueDisplayName(string displayName, ListObjectModel skipModel)
        {
            var skipIndex = _models.IndexOf(skipModel);
            ListObjectSupport.ValidateUniqueDisplayName(_models, displayName, skipIndex);
        }

        void IListObjectOwner.ValidateNoOverlap(int startRow, int startColumn, int endRow, int endColumn, ListObjectModel skipModel)
        {
            var skipIndex = _models.IndexOf(skipModel);
            ListObjectSupport.ValidateNoOverlap(_models, startRow, startColumn, endRow, endColumn, skipIndex);
        }

        void IListObjectOwner.RemoveTable(ListObjectModel model)
        {
            _models.Remove(model);
        }
    }
}
