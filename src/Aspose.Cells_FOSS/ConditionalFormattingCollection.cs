using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class ConditionalFormattingCollection
{
    private readonly List<ConditionalFormattingModel> _collections;

    internal ConditionalFormattingCollection(List<ConditionalFormattingModel> collections)
    {
        _collections = collections;
    }

    public int Count
    {
        get
        {
            return _collections.Count;
        }
    }

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

    public int Add()
    {
        _collections.Add(new ConditionalFormattingModel());
        return _collections.Count - 1;
    }

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= _collections.Count)
        {
            throw new CellsException("Conditional formatting index was out of range.");
        }

        _collections.RemoveAt(index);
    }

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
