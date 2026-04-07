using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class FormatConditionCollection
{
    private readonly List<ConditionalFormattingModel> _owner;
    private readonly ConditionalFormattingModel _model;

    internal FormatConditionCollection(List<ConditionalFormattingModel> owner, ConditionalFormattingModel model)
    {
        _owner = owner;
        _model = model;
    }

    public int Count
    {
        get
        {
            return _model.Conditions.Count;
        }
    }

    public int RangeCount
    {
        get
        {
            return _model.Areas.Count;
        }
    }

    public FormatCondition this[int index]
    {
        get
        {
            if (index < 0 || index >= _model.Conditions.Count)
            {
                throw new CellsException("Format condition index was out of range.");
            }

            return new FormatCondition(_owner, _model, _model.Conditions[index]);
        }
    }

    public int Add(CellArea area, FormatConditionType type, OperatorType operatorType, string formula1, string formula2)
    {
        AddArea(area);
        return AddCondition(type, operatorType, formula1, formula2);
    }

    public int AddCondition(FormatConditionType type)
    {
        return AddCondition(type, OperatorType.None, string.Empty, string.Empty);
    }

    public int AddCondition(FormatConditionType type, OperatorType operatorType, string formula1, string formula2)
    {
        var condition = new FormatConditionModel
        {
            Type = type,
            Operator = operatorType,
            Formula1 = NormalizeFormula(formula1),
            Formula2 = NormalizeFormula(formula2),
            Priority = ConditionalFormattingCollection.GetNextPriority(_owner),
            Style = StyleValue.Default.Clone(),
        };
        InitializeDefaults(condition);
        _model.Conditions.Add(condition);
        return _model.Conditions.Count - 1;
    }

    public void AddArea(CellArea area)
    {
        ValidateArea(area);
        _model.Areas.Add(area);
        SortAreas(_model.Areas);
    }

    public CellArea GetCellArea(int index)
    {
        if (index < 0 || index >= _model.Areas.Count)
        {
            throw new CellsException("Conditional formatting area index was out of range.");
        }

        return _model.Areas[index];
    }

    public void RemoveArea(int index)
    {
        if (index < 0 || index >= _model.Areas.Count)
        {
            throw new CellsException("Conditional formatting area index was out of range.");
        }

        _model.Areas.RemoveAt(index);
        RemoveCollectionIfEmpty(_owner, _model);
    }

    public void RemoveArea(int startRow, int startColumn, int totalRows, int totalColumns)
    {
        RemoveArea(new CellArea(startRow, startColumn, totalRows, totalColumns));
    }

    internal void RemoveArea(CellArea area)
    {
        ValidateArea(area);
        ReplaceAreas(_model, SubtractAreas(_model.Areas, area));
        RemoveCollectionIfEmpty(_owner, _model);
    }

    public void RemoveCondition(int index)
    {
        if (index < 0 || index >= _model.Conditions.Count)
        {
            throw new CellsException("Format condition index was out of range.");
        }

        _model.Conditions.RemoveAt(index);
        RemoveCollectionIfEmpty(_owner, _model);
    }

    internal static void RemoveCondition(IList<ConditionalFormattingModel> owner, ConditionalFormattingModel collection, FormatConditionModel model)
    {
        collection.Conditions.Remove(model);
        RemoveCollectionIfEmpty(owner, collection);
    }

    private static void InitializeDefaults(FormatConditionModel condition)
    {
        switch (condition.Type)
        {
            case FormatConditionType.DuplicateValues:
                condition.Duplicate = true;
                break;
            case FormatConditionType.UniqueValues:
                condition.Duplicate = false;
                break;
            case FormatConditionType.Top10:
                condition.Top = true;
                condition.Rank = 10;
                break;
            case FormatConditionType.Bottom10:
                condition.Top = false;
                condition.Rank = 10;
                break;
            case FormatConditionType.AboveAverage:
                condition.Above = true;
                break;
            case FormatConditionType.BelowAverage:
                condition.Above = false;
                break;
            case FormatConditionType.ColorScale:
                condition.ColorScaleCount = 2;
                break;
            case FormatConditionType.DataBar:
                condition.BarColor = new ColorValue(255, 99, 142, 198);
                break;
            case FormatConditionType.IconSet:
                condition.IconSetType = "3TrafficLights1";
                break;
        }
    }

    private static void RemoveCollectionIfEmpty(IList<ConditionalFormattingModel> owner, ConditionalFormattingModel collection)
    {
        if (collection.Areas.Count == 0 || collection.Conditions.Count == 0)
        {
            owner.Remove(collection);
        }
    }

    private static void ReplaceAreas(ConditionalFormattingModel model, IReadOnlyList<CellArea> areas)
    {
        model.Areas.Clear();
        for (var index = 0; index < areas.Count; index++)
        {
            model.Areas.Add(areas[index]);
        }
    }

    private static string? NormalizeFormula(string? value)
    {
        if (value is null)
        {
            return null;
        }

        var trimmed = value.Trim();
        if (trimmed.Length == 0)
        {
            return null;
        }

        if (trimmed[0] == '=')
        {
            return trimmed.Substring(1);
        }

        return trimmed;
    }

    private static void ValidateArea(CellArea area)
    {
        if (area.FirstRow < 0 || area.FirstColumn < 0 || area.TotalRows <= 0 || area.TotalColumns <= 0)
        {
            throw new CellsException("Conditional formatting area must be a positive cell range.");
        }
    }

    private static List<CellArea> SubtractAreas(IReadOnlyList<CellArea> sourceAreas, CellArea removal)
    {
        var remaining = new List<CellArea>();
        for (var index = 0; index < sourceAreas.Count; index++)
        {
            SubtractArea(sourceAreas[index], removal, remaining);
        }

        SortAreas(remaining);
        return remaining;
    }

    private static void SubtractArea(CellArea source, CellArea removal, IList<CellArea> output)
    {
        if (!AreasOverlap(source, removal))
        {
            output.Add(source);
            return;
        }

        var sourceLastRow = source.FirstRow + source.TotalRows - 1;
        var sourceLastColumn = source.FirstColumn + source.TotalColumns - 1;
        var removalLastRow = removal.FirstRow + removal.TotalRows - 1;
        var removalLastColumn = removal.FirstColumn + removal.TotalColumns - 1;

        var overlapFirstRow = Math.Max(source.FirstRow, removal.FirstRow);
        var overlapFirstColumn = Math.Max(source.FirstColumn, removal.FirstColumn);
        var overlapLastRow = Math.Min(sourceLastRow, removalLastRow);
        var overlapLastColumn = Math.Min(sourceLastColumn, removalLastColumn);

        AddIfNonEmpty(output, source.FirstRow, source.FirstColumn, overlapFirstRow - 1, sourceLastColumn);
        AddIfNonEmpty(output, overlapLastRow + 1, source.FirstColumn, sourceLastRow, sourceLastColumn);
        AddIfNonEmpty(output, overlapFirstRow, source.FirstColumn, overlapLastRow, overlapFirstColumn - 1);
        AddIfNonEmpty(output, overlapFirstRow, overlapLastColumn + 1, overlapLastRow, sourceLastColumn);
    }

    private static void AddIfNonEmpty(IList<CellArea> areas, int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        if (lastRow < firstRow || lastColumn < firstColumn)
        {
            return;
        }

        areas.Add(CellArea.CreateCellArea(firstRow, firstColumn, lastRow, lastColumn));
    }

    internal static bool AreasOverlap(CellArea left, CellArea right)
    {
        var leftLastRow = left.FirstRow + left.TotalRows - 1;
        var leftLastColumn = left.FirstColumn + left.TotalColumns - 1;
        var rightLastRow = right.FirstRow + right.TotalRows - 1;
        var rightLastColumn = right.FirstColumn + right.TotalColumns - 1;

        return left.FirstRow <= rightLastRow
            && right.FirstRow <= leftLastRow
            && left.FirstColumn <= rightLastColumn
            && right.FirstColumn <= leftLastColumn;
    }

    internal static void SortAreas(IList<CellArea> areas)
    {
        if (areas is List<CellArea> list)
        {
            list.Sort(CompareAreas);
            return;
        }

        var ordered = new List<CellArea>(areas.Count);
        for (var index = 0; index < areas.Count; index++)
        {
            ordered.Add(areas[index]);
        }

        ordered.Sort(CompareAreas);
        areas.Clear();
        for (var index = 0; index < ordered.Count; index++)
        {
            areas.Add(ordered[index]);
        }
    }

    internal static int CompareAreas(CellArea left, CellArea right)
    {
        var rowComparison = left.FirstRow.CompareTo(right.FirstRow);
        if (rowComparison != 0)
        {
            return rowComparison;
        }

        var columnComparison = left.FirstColumn.CompareTo(right.FirstColumn);
        if (columnComparison != 0)
        {
            return columnComparison;
        }

        var rowCountComparison = left.TotalRows.CompareTo(right.TotalRows);
        if (rowCountComparison != 0)
        {
            return rowCountComparison;
        }

        return left.TotalColumns.CompareTo(right.TotalColumns);
    }
}
