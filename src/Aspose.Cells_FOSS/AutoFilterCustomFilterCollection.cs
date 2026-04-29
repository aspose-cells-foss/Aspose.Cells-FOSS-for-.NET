using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a collection of auto filter custom filter objects.
    /// </summary>
    public sealed class AutoFilterCustomFilterCollection : IEnumerable<AutoFilterCustomFilter>
    {
        private readonly List<AutoFilterCustomFilterModel> _models;
        private readonly FilterColumnModel _columnModel;

        internal AutoFilterCustomFilterCollection(List<AutoFilterCustomFilterModel> models, FilterColumnModel columnModel)
        {
            _models = models;
            _columnModel = columnModel;
        }

        /// <summary>
        /// Gets or sets a value indicating whether match all.
        /// </summary>
        public bool MatchAll
        {
            get
            {
                return _columnModel.CustomFiltersAnd;
            }
            set
            {
                _columnModel.CustomFiltersAnd = value;
            }
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
        public AutoFilterCustomFilter this[int index]
        {
            get
            {
                if (index < 0 || index >= _models.Count)
                {
                    throw new CellsException("Custom filter index was out of range.");
                }

                return new AutoFilterCustomFilter(_models[index]);
            }
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <param name="operatorType">The operator type.</param>
        /// <param name="value">The value.</param>
        /// <returns>The zero-based index of the added item.</returns>
        public int Add(FilterOperatorType operatorType, string value)
        {
            if (_models.Count >= 2)
            {
                throw new CellsException("Custom filters support at most two filter conditions.");
            }

            var model = new AutoFilterCustomFilterModel
            {
                Operator = AutoFilterSupport.ToOperatorName(operatorType) ?? string.Empty,
                Value = AutoFilterSupport.NormalizeText(value, nameof(value)),
            };
            _models.Add(model);
            return _models.Count - 1;
        }

        /// <summary>
        /// Removes the specified item.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _models.Count)
            {
                throw new CellsException("Custom filter index was out of range.");
            }

            _models.RemoveAt(index);
        }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            _models.Clear();
            _columnModel.CustomFiltersAnd = false;
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator<AutoFilterCustomFilter> GetEnumerator()
        {
            var filters = new List<AutoFilterCustomFilter>(_models.Count);
            for (var index = 0; index < _models.Count; index++)
            {
                filters.Add(new AutoFilterCustomFilter(_models[index]));
            }

            return filters.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
