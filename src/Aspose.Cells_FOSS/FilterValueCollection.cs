using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a collection of filter value objects.
    /// </summary>
    public sealed class FilterValueCollection : IEnumerable<string>
    {
        private readonly List<string> _values;

        internal FilterValueCollection(List<string> values)
        {
            _values = values;
        }

        /// <summary>
        /// Gets the number of items.
        /// </summary>
        public int Count
        {
            get
            {
                return _values.Count;
            }
        }

        /// <summary>
        /// Gets the element at the specified zero-based index.
        /// </summary>
        public string this[int index]
        {
            get
            {
                if (index < 0 || index >= _values.Count)
                {
                    throw new CellsException("Filter value index was out of range.");
                }

                return _values[index];
            }
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The zero-based index of the added item.</returns>
        public int Add(string value)
        {
            var normalized = AutoFilterSupport.NormalizeText(value, nameof(value));
            _values.Add(normalized);
            return _values.Count - 1;
        }

        /// <summary>
        /// Removes the specified item.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _values.Count)
            {
                throw new CellsException("Filter value index was out of range.");
            }

            _values.RemoveAt(index);
        }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            _values.Clear();
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator<string> GetEnumerator()
        {
            var values = new List<string>(_values.Count);
            for (var index = 0; index < _values.Count; index++)
            {
                values.Add(_values[index]);
            }

            return values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
