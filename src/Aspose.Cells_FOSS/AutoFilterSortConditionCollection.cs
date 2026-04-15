using System.Linq;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a collection of auto filter sort condition objects.
    /// </summary>
    public sealed class AutoFilterSortConditionCollection : IEnumerable<AutoFilterSortCondition>
    {
        private readonly List<AutoFilterSortConditionModel> _models;

        internal AutoFilterSortConditionCollection(List<AutoFilterSortConditionModel> models)
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
        public AutoFilterSortCondition this[int index]
        {
            get
            {
                if (index < 0 || index >= _models.Count)
                {
                    throw new CellsException("Sort condition index was out of range.");
                }

                return new AutoFilterSortCondition(_models[index]);
            }
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <param name="reference">The reference.</param>
        /// <returns>The zero-based index of the added item.</returns>
        public int Add(string reference)
        {
            var model = new AutoFilterSortConditionModel
            {
                Ref = AutoFilterSupport.NormalizeRequiredRange(reference, nameof(reference)),
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
                throw new CellsException("Sort condition index was out of range.");
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
        public IEnumerator<AutoFilterSortCondition> GetEnumerator()
        {
            var conditions = new List<AutoFilterSortCondition>(_models.Count);
            for (var index = 0; index < _models.Count; index++)
            {
                conditions.Add(new AutoFilterSortCondition(_models[index]));
            }

            return conditions.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
