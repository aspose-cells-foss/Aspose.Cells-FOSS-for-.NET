using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class HyperlinkCollection
{
    private readonly List<HyperlinkModel> _hyperlinks;

    internal HyperlinkCollection(List<HyperlinkModel> hyperlinks)
    {
        _hyperlinks = hyperlinks;
    }

    public int Count
    {
        get
        {
            return _hyperlinks.Count;
        }
    }

    public Hyperlink this[int index]
    {
        get
        {
            if (index < 0 || index >= _hyperlinks.Count)
            {
                throw new CellsException("Hyperlink index was out of range.");
            }

            return new Hyperlink(_hyperlinks, _hyperlinks[index]);
        }
    }

    public int Add(string cellName, int totalRows, int totalColumns, string address)
    {
        if (address is null)
        {
            throw new ArgumentNullException(nameof(address));
        }

        return AddInternal(cellName, totalRows, totalColumns, address);
    }

    public int Add(int firstRow, int firstColumn, int totalRows, int totalColumns, string address)
    {
        if (firstRow < 0 || firstColumn < 0)
        {
            throw new CellsException("Hyperlink origin indices must be non-negative.");
        }

        if (address is null)
        {
            throw new ArgumentNullException(nameof(address));
        }

        return AddInternal(null, totalRows, totalColumns, address, firstRow, firstColumn);
    }

    public int Add(string startCellName, string endCellName, string address, string textToDisplay, string screenTip)
    {
        if (startCellName is null)
        {
            throw new ArgumentNullException(nameof(startCellName));
        }

        if (endCellName is null)
        {
            throw new ArgumentNullException(nameof(endCellName));
        }

        if (address is null)
        {
            throw new ArgumentNullException(nameof(address));
        }

        CellAddress startAddress;
        CellAddress endAddress;
        try
        {
            startAddress = CellAddress.Parse(startCellName);
            endAddress = CellAddress.Parse(endCellName);
        }
        catch (ArgumentException exception)
        {
            throw new CellsException(exception.Message, exception);
        }

        var firstRow = Math.Min(startAddress.RowIndex, endAddress.RowIndex);
        var firstColumn = Math.Min(startAddress.ColumnIndex, endAddress.ColumnIndex);
        var lastRow = Math.Max(startAddress.RowIndex, endAddress.RowIndex);
        var lastColumn = Math.Max(startAddress.ColumnIndex, endAddress.ColumnIndex);
        var index = AddInternal(null, lastRow - firstRow + 1, lastColumn - firstColumn + 1, address, firstRow, firstColumn);
        var hyperlink = _hyperlinks[index];
        hyperlink.TextToDisplay = NormalizeText(textToDisplay);
        hyperlink.ScreenTip = NormalizeText(screenTip);
        return index;
    }

    public void RemoveAt(int index)
    {
        if (index < 0 || index >= _hyperlinks.Count)
        {
            throw new CellsException("Hyperlink index was out of range.");
        }

        _hyperlinks.RemoveAt(index);
    }

    private int AddInternal(string? cellName, int totalRows, int totalColumns, string address, int? firstRowOverride = null, int? firstColumnOverride = null)
    {
        if (totalRows <= 0 || totalColumns <= 0)
        {
            throw new CellsException("Hyperlink range dimensions must be positive.");
        }

        CellAddress anchor;
        if (firstRowOverride.HasValue && firstColumnOverride.HasValue)
        {
            anchor = new CellAddress(firstRowOverride.Value, firstColumnOverride.Value);
        }
        else
        {
            if (string.IsNullOrWhiteSpace(cellName))
            {
                throw new CellsException("Hyperlink anchor must be a valid cell reference.");
            }

            try
            {
                anchor = CellAddress.Parse(cellName!);
            }
            catch (ArgumentException exception)
            {
                throw new CellsException(exception.Message, exception);
            }
        }

        var candidate = new HyperlinkModel
        {
            FirstRow = anchor.RowIndex,
            FirstColumn = anchor.ColumnIndex,
            TotalRows = totalRows,
            TotalColumns = totalColumns,
        };
        AssignAddress(candidate, address);

        for (var index = 0; index < _hyperlinks.Count; index++)
        {
            if (Overlaps(_hyperlinks[index], candidate))
            {
                throw new CellsException("Hyperlink ranges must not overlap.");
            }
        }

        _hyperlinks.Add(candidate);
        return _hyperlinks.Count - 1;
    }

    private static void AssignAddress(HyperlinkModel model, string? address)
    {
        if (string.IsNullOrWhiteSpace(address))
        {
            throw new CellsException("Hyperlink address must be non-empty.");
        }

        var normalized = address!.Trim();
        if (normalized.StartsWith("#", StringComparison.Ordinal))
        {
            model.Address = null;
            model.SubAddress = normalized.Substring(1);
            return;
        }

        if (normalized.IndexOf('!') >= 0)
        {
            model.Address = null;
            model.SubAddress = normalized;
            return;
        }

        model.Address = normalized;
        model.SubAddress = null;
    }

    private static bool Overlaps(HyperlinkModel left, HyperlinkModel right)
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

    private static string? NormalizeText(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        return value;
    }
}

