using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class ValidationCollection
{
    private readonly List<ValidationModel> _validations;

    internal ValidationCollection(List<ValidationModel> validations)
    {
        _validations = validations;
    }

    public int Count
    {
        get
        {
            return _validations.Count;
        }
    }

    public Validation this[int index]
    {
        get
        {
            if (index < 0 || index >= _validations.Count)
            {
                throw new CellsException("Validation index was out of range.");
            }

            return new Validation(_validations, _validations[index]);
        }
    }

    public int Add(CellArea area)
    {
        ValidateArea(area);
        EnsureNoOverlap(_validations, null, area);

        var validation = new ValidationModel();
        validation.Areas.Add(area);
        _validations.Add(validation);
        return _validations.Count - 1;
    }

    public Validation? GetValidationInCell(int row, int column)
    {
        if (row < 0 || column < 0)
        {
            throw new CellsException("Row and column indices must be non-negative.");
        }

        for (var index = 0; index < _validations.Count; index++)
        {
            var validation = _validations[index];
            for (var areaIndex = 0; areaIndex < validation.Areas.Count; areaIndex++)
            {
                if (Contains(validation.Areas[areaIndex], row, column))
                {
                    return new Validation(_validations, validation);
                }
            }
        }

        return null;
    }

    public void RemoveACell(int row, int column)
    {
        if (row < 0 || column < 0)
        {
            throw new CellsException("Row and column indices must be non-negative.");
        }

        RemoveArea(new CellArea(row, column, 1, 1));
    }

    public void RemoveArea(CellArea cellArea)
    {
        ValidateArea(cellArea);

        for (var index = _validations.Count - 1; index >= 0; index--)
        {
            var validation = _validations[index];
            ReplaceAreas(validation, SubtractAreas(validation.Areas, cellArea));
            if (validation.Areas.Count == 0)
            {
                _validations.RemoveAt(index);
            }
        }
    }

    internal static void AddAreaToValidation(IList<ValidationModel> owner, ValidationModel validation, CellArea area)
    {
        ValidateArea(area);
        EnsureNoOverlap(owner, validation, area);

        for (var index = 0; index < validation.Areas.Count; index++)
        {
            if (AreasOverlap(validation.Areas[index], area))
            {
                throw new CellsException("Validation areas must not overlap.");
            }
        }

        validation.Areas.Add(area);
        SortAreas(validation.Areas);
    }

    internal static void RemoveAreaFromValidation(IList<ValidationModel> owner, ValidationModel validation, CellArea area)
    {
        ValidateArea(area);
        ReplaceAreas(validation, SubtractAreas(validation.Areas, area));
        if (validation.Areas.Count == 0)
        {
            owner.Remove(validation);
        }
    }

    private static void ReplaceAreas(ValidationModel validation, IReadOnlyList<CellArea> areas)
    {
        validation.Areas.Clear();
        for (var index = 0; index < areas.Count; index++)
        {
            validation.Areas.Add(areas[index]);
        }
    }

    private static void ValidateArea(CellArea area)
    {
        if (area.FirstRow < 0 || area.FirstColumn < 0 || area.TotalRows <= 0 || area.TotalColumns <= 0)
        {
            throw new CellsException("Validation area must be a positive cell range.");
        }
    }

    private static void EnsureNoOverlap(IList<ValidationModel> validations, ValidationModel? currentValidation, CellArea candidate)
    {
        for (var validationIndex = 0; validationIndex < validations.Count; validationIndex++)
        {
            var validation = validations[validationIndex];
            if (ReferenceEquals(validation, currentValidation))
            {
                continue;
            }

            for (var areaIndex = 0; areaIndex < validation.Areas.Count; areaIndex++)
            {
                if (AreasOverlap(validation.Areas[areaIndex], candidate))
                {
                    throw new CellsException("Validation areas must not overlap.");
                }
            }
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

    private static bool Contains(CellArea area, int row, int column)
    {
        return row >= area.FirstRow
            && row < area.FirstRow + area.TotalRows
            && column >= area.FirstColumn
            && column < area.FirstColumn + area.TotalColumns;
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