using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

/// <summary>
/// Represents column.
/// </summary>
public sealed class Column
{
    private readonly Worksheet _worksheet;
    private readonly int _index;

    internal Column(Worksheet worksheet, int index)
    {
        _worksheet = worksheet;
        _index = index;
    }

    /// <summary>
    /// Gets or sets the width.
    /// </summary>
    public double? Width
    {
        get
        {
            var model = FindModel();
            return model?.Width;
        }
        set
        {
            if (value.HasValue && value.Value <= 0d)
            {
                throw new CellsException("Column width must be positive.");
            }

            var model = GetOrCreateModel();
            model.Width = value;
            Normalize(model);
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether hidden.
    /// </summary>
    public bool IsHidden
    {
        get
        {
            var model = FindModel();
            return model?.Hidden ?? false;
        }
        set
        {
            var model = GetOrCreateModel();
            model.Hidden = value;
            Normalize(model);
        }
    }

    private ColumnRangeModel? FindModel()
    {
        for (var index = _worksheet.Model.Columns.Count - 1; index >= 0; index--)
        {
            var model = _worksheet.Model.Columns[index];
            if (model.MinColumnIndex <= _index && model.MaxColumnIndex >= _index)
            {
                return model;
            }
        }

        return null;
    }

    private ColumnRangeModel GetOrCreateModel()
    {
        var existing = FindModel();
        if (existing is not null && existing.MinColumnIndex == _index && existing.MaxColumnIndex == _index)
        {
            return existing;
        }

        var created = new ColumnRangeModel
        {
            MinColumnIndex = _index,
            MaxColumnIndex = _index,
        };
        _worksheet.Model.Columns.Add(created);
        return created;
    }

    private void Normalize(ColumnRangeModel model)
    {
        if (!model.Width.HasValue && !model.Hidden && !model.StyleIndex.HasValue)
        {
            _worksheet.Model.Columns.Remove(model);
        }
    }
}
