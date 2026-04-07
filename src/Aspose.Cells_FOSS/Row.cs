using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class Row
{
    private readonly Worksheet _worksheet;
    private readonly int _index;

    internal Row(Worksheet worksheet, int index)
    {
        _worksheet = worksheet;
        _index = index;
    }

    public double? Height
    {
        get
        {
            var model = TryGetModel();
            return model?.Height;
        }
        set
        {
            if (value.HasValue && value.Value <= 0d)
            {
                throw new CellsException("Row height must be positive.");
            }

            var model = GetOrCreateModel();
            model.Height = value;
            Normalize();
        }
    }

    public bool IsHidden
    {
        get
        {
            var model = TryGetModel();
            return model?.Hidden ?? false;
        }
        set
        {
            var model = GetOrCreateModel();
            model.Hidden = value;
            Normalize();
        }
    }

    private RowModel? TryGetModel()
    {
        _worksheet.Model.Rows.TryGetValue(_index, out var model);
        return model;
    }

    private RowModel GetOrCreateModel()
    {
        if (_worksheet.Model.Rows.TryGetValue(_index, out var existing))
        {
            return existing;
        }

        var created = new RowModel();
        _worksheet.Model.Rows[_index] = created;
        return created;
    }

    private void Normalize()
    {
        if (_worksheet.Model.Rows.TryGetValue(_index, out var model)
            && !model.Height.HasValue
            && !model.Hidden
            && !model.StyleIndex.HasValue)
        {
            _worksheet.Model.Rows.Remove(_index);
        }
    }
}