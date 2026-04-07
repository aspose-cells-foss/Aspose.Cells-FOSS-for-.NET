namespace Aspose.Cells_FOSS.Core;

public sealed class WorksheetProtectionModel
{
    public bool IsProtected { get; set; }
    public bool Objects { get; set; }
    public bool Scenarios { get; set; }
    public bool FormatCells { get; set; }
    public bool FormatColumns { get; set; }
    public bool FormatRows { get; set; }
    public bool InsertColumns { get; set; }
    public bool InsertRows { get; set; }
    public bool InsertHyperlinks { get; set; }
    public bool DeleteColumns { get; set; }
    public bool DeleteRows { get; set; }
    public bool SelectLockedCells { get; set; }
    public bool Sort { get; set; }
    public bool AutoFilter { get; set; }
    public bool PivotTables { get; set; }
    public bool SelectUnlockedCells { get; set; }
    public string? PasswordHash { get; set; }
    public string? AlgorithmName { get; set; }
    public string? HashValue { get; set; }
    public string? SaltValue { get; set; }
    public string? SpinCount { get; set; }

    public void Clear()
    {
        IsProtected = false;
        Objects = false;
        Scenarios = false;
        FormatCells = false;
        FormatColumns = false;
        FormatRows = false;
        InsertColumns = false;
        InsertRows = false;
        InsertHyperlinks = false;
        DeleteColumns = false;
        DeleteRows = false;
        SelectLockedCells = false;
        Sort = false;
        AutoFilter = false;
        PivotTables = false;
        SelectUnlockedCells = false;
        PasswordHash = null;
        AlgorithmName = null;
        HashValue = null;
        SaltValue = null;
        SpinCount = null;
    }

    public bool HasStoredState()
    {
        return IsProtected
            || Objects
            || Scenarios
            || FormatCells
            || FormatColumns
            || FormatRows
            || InsertColumns
            || InsertRows
            || InsertHyperlinks
            || DeleteColumns
            || DeleteRows
            || SelectLockedCells
            || Sort
            || AutoFilter
            || PivotTables
            || SelectUnlockedCells
            || !string.IsNullOrWhiteSpace(PasswordHash)
            || !string.IsNullOrWhiteSpace(AlgorithmName)
            || !string.IsNullOrWhiteSpace(HashValue)
            || !string.IsNullOrWhiteSpace(SaltValue)
            || !string.IsNullOrWhiteSpace(SpinCount);
    }
}
