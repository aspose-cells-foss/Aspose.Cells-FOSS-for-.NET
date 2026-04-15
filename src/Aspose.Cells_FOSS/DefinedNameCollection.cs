using System.Linq;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a collection of defined name objects.
    /// </summary>
    public sealed class DefinedNameCollection : IEnumerable<DefinedName>
    {
        private readonly Workbook _workbook;

        internal DefinedNameCollection(Workbook workbook)
        {
            _workbook = workbook;
        }

        /// <summary>
        /// Gets the number of items.
        /// </summary>
        public int Count
        {
            get
            {
                return _workbook.Model.DefinedNames.Count;
            }
        }

        /// <summary>
        /// Gets the element at the specified zero-based index.
        /// </summary>
        public DefinedName this[int index]
        {
            get
            {
                if (index < 0 || index >= _workbook.Model.DefinedNames.Count)
                {
                    throw new CellsException("Defined name index was out of range.");
                }

                return new DefinedName(_workbook, _workbook.Model.DefinedNames[index]);
            }
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="formula">The formula.</param>
        /// <returns>The zero-based index of the added item.</returns>
        public int Add(string name, string formula)
        {
            return Add(name, formula, null);
        }

        /// <summary>
        /// Adds the specified item.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="formula">The formula.</param>
        /// <param name="localSheetIndex">The local sheet index.</param>
        /// <returns>The zero-based index of the added item.</returns>
        public int Add(string name, string formula, int? localSheetIndex)
        {
            var normalizedName = DefinedNameUtility.NormalizeName(name);
            var normalizedFormula = DefinedNameUtility.NormalizeFormula(formula);
            _workbook.EnsureValidDefinedNameScope(localSheetIndex);
            _workbook.EnsureUniqueDefinedName(null, normalizedName, localSheetIndex);

            var model = new DefinedNameModel
            {
                Name = normalizedName,
                Formula = normalizedFormula,
                LocalSheetIndex = localSheetIndex,
            };

            _workbook.Model.DefinedNames.Add(model);
            return _workbook.Model.DefinedNames.Count - 1;
        }

        /// <summary>
        /// Removes the specified item.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _workbook.Model.DefinedNames.Count)
            {
                throw new CellsException("Defined name index was out of range.");
            }

            _workbook.Model.DefinedNames.RemoveAt(index);
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator<DefinedName> GetEnumerator()
        {
            var names = new List<DefinedName>(_workbook.Model.DefinedNames.Count);
            for (var index = 0; index < _workbook.Model.DefinedNames.Count; index++)
            {
                names.Add(new DefinedName(_workbook, _workbook.Model.DefinedNames[index]));
            }

            return names.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
