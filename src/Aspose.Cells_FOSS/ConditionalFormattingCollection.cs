using System.IO;
using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a collection of conditional formatting objects.
    /// </summary>
    public sealed class ConditionalFormattingCollection
    {
        private readonly List<ConditionalFormattingModel> _collections;

        internal ConditionalFormattingCollection(List<ConditionalFormattingModel> collections)
        {
            _collections = collections;
        }

        /// <summary>
        /// Gets the number of items.
        /// </summary>
        public int Count
        {
            get
            {
                return _collections.Count;
            }
        }

        /// <summary>
        /// Gets the element at the specified zero-based index.
        /// </summary>
        public FormatConditionCollection this[int index]
        {
            get
            {
                if (index < 0 || index >= _collections.Count)
                {
                    throw new CellsException("Conditional formatting index was out of range.");
                }

                return new FormatConditionCollection(_collections, _collections[index]);
            }
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <returns>The zero-based index of the added item.</returns>
        public int Add()
        {
            _collections.Add(new ConditionalFormattingModel());
            return _collections.Count - 1;
        }

        /// <summary>
        /// Removes the specified item.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _collections.Count)
            {
                throw new CellsException("Conditional formatting index was out of range.");
            }

            _collections.RemoveAt(index);
        }

        /// <summary>
        /// Removes the specified item.
        /// </summary>
        /// <param name="startRow">The start row.</param>
        /// <param name="startColumn">The start column.</param>
        /// <param name="totalRows">The total number of rows.</param>
        /// <param name="totalColumns">The total number of columns.</param>
        public void RemoveArea(int startRow, int startColumn, int totalRows, int totalColumns)
        {
            var area = new CellArea(startRow, startColumn, totalRows, totalColumns);
            for (var index = _collections.Count - 1; index >= 0; index--)
            {
                var collection = new FormatConditionCollection(_collections, _collections[index]);
                collection.RemoveArea(area);
            }
        }

        internal static int GetNextPriority(IReadOnlyList<ConditionalFormattingModel> collections)
        {
            var maxPriority = 0;
            for (var collectionIndex = 0; collectionIndex < collections.Count; collectionIndex++)
            {
                var collection = collections[collectionIndex];
                for (var conditionIndex = 0; conditionIndex < collection.Conditions.Count; conditionIndex++)
                {
                    maxPriority = Math.Max(maxPriority, collection.Conditions[conditionIndex].Priority);
                }
            }

            return maxPriority + 1;
        }
    }
}
