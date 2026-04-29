using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a collection of filter column objects.
    /// </summary>
    public sealed class FilterColumnCollection : IEnumerable<FilterColumn>
    {
        private readonly List<FilterColumnModel> _models;

        internal FilterColumnCollection(List<FilterColumnModel> models)
        {
            _models = models;
        }

        /// <summary>
        /// Gets the number of items.
        /// </summary>
        public int Count
        {
            get
            {
                return _models.Count;
            }
        }

        /// <summary>
        /// Gets the element at the specified zero-based index.
        /// </summary>
        public FilterColumn this[int index]
        {
            get
            {
                if (index < 0 || index >= _models.Count)
                {
                    throw new CellsException("Filter column index was out of range.");
                }

                return new FilterColumn(_models[index]);
            }
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <param name="columnIndex">The zero-based column index.</param>
        /// <returns>The zero-based index of the added item.</returns>
        public int Add(int columnIndex)
        {
            if (columnIndex < 0)
            {
                throw new CellsException("Filter column index must be zero or greater.");
            }

            for (var index = 0; index < _models.Count; index++)
            {
                if (_models[index].ColumnIndex == columnIndex)
                {
                    throw new CellsException("A filter column for the specified column index already exists.");
                }
            }

            var model = new FilterColumnModel();
            model.ColumnIndex = columnIndex;

            var insertIndex = 0;
            while (insertIndex < _models.Count && _models[insertIndex].ColumnIndex < columnIndex)
            {
                insertIndex++;
            }

            _models.Insert(insertIndex, model);
            return insertIndex;
        }

        /// <summary>
        /// Removes the specified item.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _models.Count)
            {
                throw new CellsException("Filter column index was out of range.");
            }

            _models.RemoveAt(index);
        }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            _models.Clear();
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator<FilterColumn> GetEnumerator()
        {
            var columns = new List<FilterColumn>(_models.Count);
            for (var index = 0; index < _models.Count; index++)
            {
                columns.Add(new FilterColumn(_models[index]));
            }

            return columns.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
