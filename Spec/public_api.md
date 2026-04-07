# Aspose.Cells FOSS for .NET - Public API Design (v0.1)

This document defines the initial public object model.

The API should remain as close as practical to Aspose.Cells for .NET for
supported v0.1 features.

## Workbook

`````csharp
public class Workbook : IDisposable
{
    public Workbook();
    public Workbook(string fileName);
    public Workbook(Stream stream);
    public Workbook(string fileName, LoadOptions options);
    public Workbook(Stream stream, LoadOptions options);

    public WorksheetCollection Worksheets { get; }
    public WorkbookSettings Settings { get; }
    public WorkbookProperties Properties { get; }
    public DocumentProperties DocumentProperties { get; }
    public DefinedNameCollection DefinedNames { get; }
    public LoadDiagnostics LoadDiagnostics { get; }

    public void Save(string fileName);
    public void Save(string fileName, SaveFormat format);
    public void Save(string fileName, SaveOptions options);
    public void Save(Stream stream, SaveFormat format);
    public void Save(Stream stream, SaveOptions options);
    public void Dispose();
}
`````

## WorkbookProperties

`````csharp
public sealed class WorkbookProperties
{
    public string CodeName { get; set; }
    public string ShowObjects { get; set; }
    public bool FilterPrivacy { get; set; }
    public bool ShowBorderUnselectedTables { get; set; }
    public bool ShowInkAnnotation { get; set; }
    public bool BackupFile { get; set; }
    public bool SaveExternalLinkValues { get; set; }
    public string UpdateLinks { get; set; }
    public bool HidePivotFieldList { get; set; }
    public int? DefaultThemeVersion { get; set; }
    public WorkbookProtection Protection { get; }
    public WorkbookView View { get; }
    public CalculationProperties Calculation { get; }
}
`````

## WorkbookProtection

`````csharp
public sealed class WorkbookProtection
{
    public bool LockStructure { get; set; }
    public bool LockWindows { get; set; }
    public bool LockRevision { get; set; }
    public string WorkbookPassword { get; set; }
    public string RevisionsPassword { get; set; }
    public bool IsProtected { get; }
}
`````

## WorkbookView

`````csharp
public sealed class WorkbookView
{
    public int XWindow { get; set; }
    public int YWindow { get; set; }
    public int WindowWidth { get; set; }
    public int WindowHeight { get; set; }
    public int ActiveTab { get; set; }
    public int FirstSheet { get; set; }
    public bool ShowHorizontalScroll { get; set; }
    public bool ShowVerticalScroll { get; set; }
    public bool ShowSheetTabs { get; set; }
    public int TabRatio { get; set; }
    public string Visibility { get; set; }
    public bool Minimized { get; set; }
    public bool AutoFilterDateGrouping { get; set; }
}
`````

## CalculationProperties

`````csharp
public sealed class CalculationProperties
{
    public int? CalculationId { get; set; }
    public string CalculationMode { get; set; }
    public bool FullCalculationOnLoad { get; set; }
    public string ReferenceMode { get; set; }
    public bool Iterate { get; set; }
    public int IterateCount { get; set; }
    public double IterateDelta { get; set; }
    public bool FullPrecision { get; set; }
    public bool CalculationCompleted { get; set; }
    public bool CalculationOnSave { get; set; }
    public bool ConcurrentCalculation { get; set; }
    public bool ForceFullCalculation { get; set; }
}
`````

## DocumentProperties

`````csharp
public sealed class DocumentProperties
{
    public CoreDocumentProperties Core { get; }
    public ExtendedDocumentProperties Extended { get; }
    public string Title { get; set; }
    public string Subject { get; set; }
    public string Author { get; set; }
    public string Keywords { get; set; }
    public string Comments { get; set; }
    public string Category { get; set; }
    public string Company { get; set; }
    public string Manager { get; set; }
}
`````

## CoreDocumentProperties

`````csharp
public sealed class CoreDocumentProperties
{
    public string Title { get; set; }
    public string Subject { get; set; }
    public string Creator { get; set; }
    public string Keywords { get; set; }
    public string Description { get; set; }
    public string LastModifiedBy { get; set; }
    public string Revision { get; set; }
    public string Category { get; set; }
    public string ContentStatus { get; set; }
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }
}
`````

## ExtendedDocumentProperties

`````csharp
public sealed class ExtendedDocumentProperties
{
    public string Application { get; set; }
    public string AppVersion { get; set; }
    public string Company { get; set; }
    public string Manager { get; set; }
    public int DocSecurity { get; set; }
    public string HyperlinkBase { get; set; }
    public bool ScaleCrop { get; set; }
    public bool LinksUpToDate { get; set; }
    public bool SharedDoc { get; set; }
}
`````

## WorksheetCollection

`````csharp
public class WorksheetCollection
{
    public Worksheet this[int index] { get; }
    public Worksheet this[string name] { get; }

    public int Count { get; }
    public int ActiveSheetIndex { get; set; }
    public string ActiveSheetName { get; set; }

    public int Add();
    public int Add(string sheetName);
    public void RemoveAt(string sheetName);
    public void RemoveAt(int index);
}
```

## Worksheet

`````csharp
public class Worksheet
{
    public string Name { get; set; }
    public VisibilityType VisibilityType { get; set; }
    public Color TabColor { get; set; }
    public bool ShowGridlines { get; set; }
    public bool ShowRowColumnHeaders { get; set; }
    public bool ShowZeros { get; set; }
    public bool RightToLeft { get; set; }
    public int Zoom { get; set; }
    public WorksheetProtection Protection { get; }
    public AutoFilter AutoFilter { get; }
    public Cells Cells { get; }
    public HyperlinkCollection Hyperlinks { get; }
    public ValidationCollection Validations { get; }
    public ConditionalFormattingCollection ConditionalFormattings { get; }
    public PageSetup PageSetup { get; }

    public void Protect();
    public void Unprotect();
}
```

## WorksheetProtection

`````csharp
public sealed class WorksheetProtection
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
}
```

## AutoFilter

`````csharp
public sealed class AutoFilter
{
    public string Range { get; set; }
    public FilterColumnCollection FilterColumns { get; }
    public AutoFilterSortState SortState { get; }

    public void Clear();
}
```

## FilterColumnCollection

`````csharp
public sealed class FilterColumnCollection
{
    public int Count { get; }
    public FilterColumn this[int index] { get; }

    public int Add(int columnIndex);
    public void RemoveAt(int index);
    public void Clear();
}
```

## FilterColumn

`````csharp
public sealed class FilterColumn
{
    public int ColumnIndex { get; }
    public bool HiddenButton { get; set; }
    public FilterValueCollection Filters { get; }
    public AutoFilterCustomFilterCollection CustomFilters { get; }
    public AutoFilterColorFilter ColorFilter { get; }
    public AutoFilterDynamicFilter DynamicFilter { get; }
    public AutoFilterTop10 Top10 { get; }

    public void Clear();
}
```

## FilterValueCollection

`````csharp
public sealed class FilterValueCollection
{
    public int Count { get; }
    public string this[int index] { get; }

    public int Add(string value);
    public void RemoveAt(int index);
    public void Clear();
}
```

## FilterOperatorType

`````csharp
public enum FilterOperatorType
{
    Equal,
    LessThan,
    LessOrEqual,
    NotEqual,
    GreaterOrEqual,
    GreaterThan,
}
```

## AutoFilterCustomFilterCollection

`````csharp
public sealed class AutoFilterCustomFilterCollection
{
    public bool MatchAll { get; set; }
    public int Count { get; }
    public AutoFilterCustomFilter this[int index] { get; }

    public int Add(FilterOperatorType operatorType, string value);
    public void RemoveAt(int index);
    public void Clear();
}
```

## AutoFilterCustomFilter

`````csharp
public sealed class AutoFilterCustomFilter
{
    public FilterOperatorType Operator { get; set; }
    public string Value { get; set; }
}
```

## AutoFilterColorFilter

`````csharp
public sealed class AutoFilterColorFilter
{
    public bool Enabled { get; set; }
    public int? DifferentialStyleId { get; set; }
    public bool CellColor { get; set; }

    public void Clear();
}
```

## AutoFilterDynamicFilter

`````csharp
public sealed class AutoFilterDynamicFilter
{
    public bool Enabled { get; set; }
    public string Type { get; set; }
    public double? Value { get; set; }
    public double? MaxValue { get; set; }

    public void Clear();
}
```

## AutoFilterTop10

`````csharp
public sealed class AutoFilterTop10
{
    public bool Enabled { get; set; }
    public bool Top { get; set; }
    public bool Percent { get; set; }
    public double? Value { get; set; }
    public double? FilterValue { get; set; }

    public void Clear();
}
```

## AutoFilterSortState

`````csharp
public sealed class AutoFilterSortState
{
    public bool ColumnSort { get; set; }
    public bool CaseSensitive { get; set; }
    public string SortMethod { get; set; }
    public string Ref { get; set; }
    public AutoFilterSortConditionCollection SortConditions { get; }

    public void Clear();
}
```

## AutoFilterSortConditionCollection

`````csharp
public sealed class AutoFilterSortConditionCollection
{
    public int Count { get; }
    public AutoFilterSortCondition this[int index] { get; }

    public int Add(string reference);
    public void RemoveAt(int index);
    public void Clear();
}
```

## AutoFilterSortCondition

`````csharp
public sealed class AutoFilterSortCondition
{
    public string Ref { get; set; }
    public bool Descending { get; set; }
    public string SortBy { get; set; }
    public string CustomList { get; set; }
    public int? DifferentialStyleId { get; set; }
    public string IconSet { get; set; }
    public int? IconId { get; set; }
}
```

## Cells
## Cells

```csharp
public class Cells
{
    public Cell this[string cellName] { get; }
    public Cell this[int row, int column] { get; }

    public void Merge(int firstRow, int firstColumn, int totalRows, int totalColumns);
}
```

## HyperlinkCollection

`````csharp
public sealed class HyperlinkCollection
{
    public int Count { get; }
    public Hyperlink this[int index] { get; }

    public int Add(string cellName, int totalRows, int totalColumns, string address);
    public int Add(int firstRow, int firstColumn, int totalRows, int totalColumns, string address);
    public int Add(string startCellName, string endCellName, string address, string textToDisplay, string screenTip);
    public void RemoveAt(int index);
}
```

## Hyperlink

`````csharp
public sealed class Hyperlink
{
    public string Area { get; }
    public string Address { get; set; }
    public TargetModeType LinkType { get; }
    public string ScreenTip { get; set; }
    public string TextToDisplay { get; set; }
    public void Delete();
}
```

## ValidationCollection

`````csharp
public sealed class ValidationCollection
{
    public int Count { get; }
    public Validation this[int index] { get; }

    public int Add(CellArea area);
    public Validation? GetValidationInCell(int row, int column);
    public void RemoveACell(int row, int column);
    public void RemoveArea(CellArea cellArea);
}
```

## Validation

`````csharp
public sealed class Validation
{
    public IReadOnlyList<CellArea> Areas { get; }
    public ValidationType Type { get; set; }
    public ValidationAlertType AlertStyle { get; set; }
    public OperatorType Operator { get; set; }
    public string Formula1 { get; set; }
    public string Formula2 { get; set; }
    public bool IgnoreBlank { get; set; }
    public bool InCellDropDown { get; set; }
    public string InputTitle { get; set; }
    public string InputMessage { get; set; }
    public string ErrorTitle { get; set; }
    public string ErrorMessage { get; set; }
    public bool ShowInput { get; set; }
    public bool ShowError { get; set; }

    public void AddArea(CellArea area);
    public void RemoveArea(CellArea area);
}
```

## ConditionalFormattingCollection

`````csharp
public sealed class ConditionalFormattingCollection
{
    public int Count { get; }
    public FormatConditionCollection this[int index] { get; }

    public int Add();
    public void RemoveAt(int index);
    public void RemoveArea(int startRow, int startColumn, int totalRows, int totalColumns);
}
```

## FormatConditionCollection

`````csharp
public sealed class FormatConditionCollection
{
    public int Count { get; }
    public int RangeCount { get; }
    public FormatCondition this[int index] { get; }

    public int Add(CellArea area, FormatConditionType type, OperatorType operatorType, string formula1, string formula2);
    public int AddCondition(FormatConditionType type);
    public int AddCondition(FormatConditionType type, OperatorType operatorType, string formula1, string formula2);
    public void AddArea(CellArea area);
    public CellArea GetCellArea(int index);
    public void RemoveArea(int index);
    public void RemoveArea(int startRow, int startColumn, int totalRows, int totalColumns);
    public void RemoveCondition(int index);
}
```

## FormatCondition

`````csharp
public sealed class FormatCondition
{
    public FormatConditionType Type { get; set; }
    public OperatorType Operator { get; set; }
    public string Formula1 { get; set; }
    public string Formula2 { get; set; }
    public string Formula { get; set; }
    public string TimePeriod { get; set; }
    public bool Duplicate { get; set; }
    public bool Top { get; set; }
    public bool Percent { get; set; }
    public int Rank { get; set; }
    public bool Above { get; set; }
    public int StandardDeviation { get; set; }
    public int ColorScaleCount { get; set; }
    public Color MinColor { get; set; }
    public Color MidColor { get; set; }
    public Color MaxColor { get; set; }
    public Color BarColor { get; set; }
    public Color NegativeBarColor { get; set; }
    public bool ShowBorder { get; set; }
    public string Direction { get; set; }
    public string BarLength { get; set; }
    public string IconSetType { get; set; }
    public bool ReverseIcons { get; set; }
    public bool ShowIconOnly { get; set; }
    public int Priority { get; set; }
    public bool StopIfTrue { get; set; }
    public Style Style { get; set; }

    public void Remove();
}
```

## FormatConditionType

`````csharp
public enum FormatConditionType
{
    CellValue,
    Expression,
    ContainsText,
    NotContainsText,
    BeginsWith,
    EndsWith,
    TimePeriod,
    DuplicateValues,
    UniqueValues,
    Top10,
    Bottom10,
    AboveAverage,
    BelowAverage,
    ColorScale,
    DataBar,
    IconSet,
}
```

## Cell

StringValue is the stable textual form of the logical cell value. DisplayStringValue applies the cell style's display formatting for supported number and date cases, using Workbook.Settings.Culture with per-format locale directives when present.

```csharp
public class Cell
{
    public object? Value { get; set; }
    public string StringValue { get; }
    public string DisplayStringValue { get; }
    public string Formula { get; set; }
    public CellValueType Type { get; }

    public void PutValue(string value);
    public void PutValue(int value);
    public void PutValue(double value);
    public void PutValue(bool value);
    public void PutValue(DateTime value);

    public Style GetStyle();
    public void SetStyle(Style style);
}
```

## Style

The `Style` object exposes a broad interface similar to Aspose.Cells.

Core persisted properties in v0.1:

- Font: `Name`, `Size`, `Bold`, `Italic`, `Underline`, `StrikeThrough`, `Color`
- Fill: `Pattern`, `ForegroundColor`, `BackgroundColor`, including classic SpreadsheetML pattern fills such as `Solid`, `LightGrid`, `DarkTrellis`, `Gray125`, and related variants
- Border: `Left`, `Right`, `Top`, `Bottom`, `Diagonal`, `DiagonalUp`, `DiagonalDown`, full classic border styles, and per-side `Color`
- Number format: `Number`, `Custom`, `NumberFormat`, and built-in/custom format lookup through `NumberFormat`
- Alignment: `HorizontalAlignment`, `VerticalAlignment`, `WrapText`, `IndentLevel`, `TextRotation`, `ShrinkToFit`, `ReadingOrder`, `RelativeIndent`
- Protection: `Locked`, `Hidden`
- Conditional-formatting differential style subset through `dxfs`

Future properties may be added without breaking compatibility.

`````csharp
public class Style
{
    public Font Font { get; set; }
    public Borders Borders { get; set; }
    public FillPattern Pattern { get; set; }
    public Color ForegroundColor { get; set; }
    public Color BackgroundColor { get; set; }
    public int Number { get; set; }
    public string? Custom { get; set; }
    public string NumberFormat { get; set; }
    public HorizontalAlignmentType HorizontalAlignment { get; set; }
    public VerticalAlignmentType VerticalAlignment { get; set; }
    public bool WrapText { get; set; }
    public int IndentLevel { get; set; }
    public int TextRotation { get; set; }
    public bool ShrinkToFit { get; set; }
    public int ReadingOrder { get; set; }
    public int RelativeIndent { get; set; }
    public bool IsLocked { get; set; }
    public bool IsHidden { get; set; }
}
```

`````csharp
public static class NumberFormat
{
    public static string GetBuiltInFormat(int formatId);
    public static bool IsBuiltInFormat(string formatCode);
    public static int? GetBuiltInFormatId(string? formatCode);
}
```
## WorkbookSettings

`````csharp
public sealed class WorkbookSettings
{
    public bool Date1904 { get; set; }
    public CultureInfo Culture { get; set; }
}
```

## DefinedNameCollection

`````csharp
public sealed class DefinedNameCollection
{
    public int Count { get; }
    public DefinedName this[int index] { get; }

    public int Add(string name, string formula);
    public int Add(string name, string formula, int? localSheetIndex);
    public void RemoveAt(int index);
}
```

## DefinedName

`````csharp
public sealed class DefinedName
{
    public string Name { get; set; }
    public string Formula { get; set; }
    public int? LocalSheetIndex { get; set; }
    public bool Hidden { get; set; }
    public string Comment { get; set; }
}
```
## PageSetup

For page margins, the Aspose-compatible public API should expose centimeter
properties such as `LeftMargin`, `RightMargin`, `TopMargin`, `BottomMargin`,
`HeaderMargin`, and `FooterMargin`, with inch-based variants exposed as
`LeftMarginInch`, `RightMarginInch`, `TopMarginInch`, `BottomMarginInch`,
`HeaderMarginInch`, and `FooterMarginInch`.














