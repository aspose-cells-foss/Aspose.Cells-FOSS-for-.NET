using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public class WorksheetCollection : IEnumerable<Worksheet>
{
    private readonly Workbook _workbook;

    internal WorksheetCollection(Workbook workbook)
    {
        _workbook = workbook;
    }

    public Worksheet this[int index]
    {
        get
        {
            return Wrap(_workbook.Model.Worksheets[index]);
        }
    }

    public Worksheet this[string name]
    {
        get
        {
            var index = IndexOf(name);
            if (index < 0)
            {
                throw new CellsException($"Worksheet '{name}' was not found.");
            }

            return Wrap(_workbook.Model.Worksheets[index]);
        }
    }

    public int Count
    {
        get
        {
            return _workbook.Model.Worksheets.Count;
        }
    }

    public int ActiveSheetIndex
    {
        get
        {
            return _workbook.Model.ActiveSheetIndex;
        }
        set
        {
            if (value < 0 || value >= _workbook.Model.Worksheets.Count)
            {
                throw new CellsException("ActiveSheetIndex must refer to an existing worksheet.");
            }

            _workbook.Model.ActiveSheetIndex = value;
        }
    }

    public string ActiveSheetName
    {
        get
        {
            return _workbook.Model.Worksheets[ActiveSheetIndex].Name;
        }
        set
        {
            var index = IndexOf(value);
            if (index < 0)
            {
                throw new CellsException($"Worksheet '{value}' was not found.");
            }

            ActiveSheetIndex = index;
        }
    }

    public int Add()
    {
        return Add(GenerateDefaultSheetName());
    }

    public int Add(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) throw new CellsException("Worksheet name must be non-empty.");
        _workbook.EnsureUniqueSheetName(sheetName);
        _workbook.Model.Worksheets.Add(new WorksheetModel(sheetName));
        return _workbook.Model.Worksheets.Count - 1;
    }

    public void RemoveAt(string sheetName)
    {
        var index = IndexOf(sheetName);
        if (index < 0)
        {
            throw new CellsException($"Worksheet '{sheetName}' was not found.");
        }

        RemoveAt(index);
    }

    public void RemoveAt(int index)
    {
        if (_workbook.Model.Worksheets.Count == 1)
        {
            throw new CellsException("A workbook must contain at least one worksheet.");
        }

        _workbook.Model.Worksheets.RemoveAt(index);
        if (_workbook.Model.ActiveSheetIndex > index)
        {
            _workbook.Model.ActiveSheetIndex--;
        }
        else if (_workbook.Model.ActiveSheetIndex >= _workbook.Model.Worksheets.Count)
        {
            _workbook.Model.ActiveSheetIndex = _workbook.Model.Worksheets.Count - 1;
        }

        var firstSheet = _workbook.Model.Properties.View.FirstSheet;
        if (firstSheet.HasValue)
        {
            if (_workbook.Model.Worksheets.Count == 0)
            {
                _workbook.Model.Properties.View.FirstSheet = 0;
            }
            else if (firstSheet.Value >= _workbook.Model.Worksheets.Count)
            {
                _workbook.Model.Properties.View.FirstSheet = _workbook.Model.Worksheets.Count - 1;
            }
        }
    }

    public IEnumerator<Worksheet> GetEnumerator()
    {
        var worksheets = new List<Worksheet>(_workbook.Model.Worksheets.Count);
        for (var i = 0; i < _workbook.Model.Worksheets.Count; i++)
        {
            worksheets.Add(Wrap(_workbook.Model.Worksheets[i]));
        }

        return worksheets.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private int IndexOf(string? name)
    {
        if (name is null) throw new ArgumentNullException(nameof(name));
        for (var i = 0; i < _workbook.Model.Worksheets.Count; i++)
        {
            var worksheet = _workbook.Model.Worksheets[i];
            if (string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }

    private string GenerateDefaultSheetName()
    {
        var suffix = 1;
        while (true)
        {
            var candidate = "Sheet" + suffix.ToString();
            if (IndexOf(candidate) < 0)
            {
                return candidate;
            }

            suffix++;
        }
    }

    private Worksheet Wrap(WorksheetModel worksheet)
    {
        return new Worksheet(_workbook, worksheet);
    }
}

