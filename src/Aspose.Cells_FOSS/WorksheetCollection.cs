using System.Collections;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

/// <summary>
/// Encapsulates the workbook's worksheets and active-sheet state.
/// </summary>
/// <example>
/// <code>
/// var workbook = new Workbook();
/// int summaryIndex = workbook.Worksheets.Add("Summary");
/// workbook.Worksheets[summaryIndex].Cells["A1"].PutValue("Ready");
/// workbook.Worksheets.ActiveSheetName = "Summary";
/// </code>
/// </example>
public class WorksheetCollection : IEnumerable<Worksheet>
{
    private readonly Workbook _workbook;

    internal WorksheetCollection(Workbook workbook)
    {
        _workbook = workbook;
    }

    /// <summary>
    /// Gets the worksheet at the specified zero-based index.
    /// </summary>
    public Worksheet this[int index]
    {
        get
        {
            return Wrap(_workbook.Model.Worksheets[index]);
        }
    }

    /// <summary>
    /// Gets the worksheet with the specified name using case-insensitive lookup.
    /// </summary>
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

    /// <summary>
    /// Gets the total number of worksheets in the workbook.
    /// </summary>
    public int Count
    {
        get
        {
            return _workbook.Model.Worksheets.Count;
        }
    }

    /// <summary>
    /// Gets or sets the zero-based index of the active worksheet.
    /// </summary>
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

    /// <summary>
    /// Gets or sets the name of the active worksheet.
    /// </summary>
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

    /// <summary>
    /// Adds a worksheet using the next generated default name.
    /// </summary>
    public int Add()
    {
        return Add(GenerateDefaultSheetName());
    }

    /// <summary>
    /// Adds a worksheet with the specified name and returns its index.
    /// </summary>
    public int Add(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) throw new CellsException("Worksheet name must be non-empty.");
        _workbook.EnsureUniqueSheetName(sheetName);
        _workbook.Model.Worksheets.Add(new WorksheetModel(sheetName));
        return _workbook.Model.Worksheets.Count - 1;
    }

    /// <summary>
    /// Removes the worksheet with the specified name.
    /// </summary>
    public void RemoveAt(string sheetName)
    {
        var index = IndexOf(sheetName);
        if (index < 0)
        {
            throw new CellsException($"Worksheet '{sheetName}' was not found.");
        }

        RemoveAt(index);
    }

    /// <summary>
    /// Removes the worksheet at the specified zero-based index.
    /// </summary>
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

    /// <summary>
    /// Returns an enumerator that iterates through worksheets in workbook order.
    /// </summary>
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
