using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class DefinedNameCollection : IEnumerable<DefinedName>
{
    private readonly Workbook _workbook;

    internal DefinedNameCollection(Workbook workbook)
    {
        _workbook = workbook;
    }

    public int Count
    {
        get
        {
            return _workbook.Model.DefinedNames.Count;
        }
    }

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

    public int Add(string name, string formula)
    {
        return Add(name, formula, null);
    }

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

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= _workbook.Model.DefinedNames.Count)
        {
            throw new CellsException("Defined name index was out of range.");
        }

        _workbook.Model.DefinedNames.RemoveAt(index);
    }

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
