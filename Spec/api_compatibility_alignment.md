# Public API Compatibility Alignment Plan

## Goal

Bring the supported public API surface closer to Aspose.Cells for .NET for
the existing v0.1 feature scope.

This plan focuses on source compatibility first, then on behavior alignment.

## Scope

Included in this alignment pass:

- `Workbook`
- `WorkbookSettings`
- `WorksheetCollection`
- `Worksheet`
- `Cell`
- `HyperlinkCollection`
- `Hyperlink`
- `PageSetup`

Excluded from this pass:

- unsupported product-scope features
- full Aspose.Cells option hierarchy replication
- locale-sensitive display formatting parity
- API surfaces for charts, drawings, tables, and encryption

## Compatibility Rules

1. Prefer the Aspose.Cells member name and signature when the project already
   supports the underlying feature.
2. If a previous FOSS-specific signature conflicts with Aspose.Cells and both
   cannot coexist, switch to the Aspose-compatible form and update tests.
3. Remove temporary bridge members once the Aspose-compatible shape exists.
4. Preserve deterministic XLSX behavior and recovery behavior while changing
   the public surface.

## Phase 1

### Cell

- Make `Cell.Value` writable.
- Add `Cell.DisplayStringValue`.
- Keep `PutValue(...)` overloads as the explicit typed entry points.
- Route value assignment through supported scalar type normalization.

### Worksheet and workbook settings

- Add `Worksheet.VisibilityType`.
- Add `WorkbookSettings.Date1904`.
- Remove legacy worksheet visibility and date-system bridge members.

### Worksheet collection

- Add `WorksheetCollection.Count`.
- Add `WorksheetCollection.ActiveSheetIndex`.
- Add `WorksheetCollection.ActiveSheetName`.
- Add parameterless `Add()`.
- Add `RemoveAt(string)`.
- Persist active sheet index through workbook XML where practical.

## Phase 2

### Hyperlinks

- Make `HyperlinkCollection.Add(...)` return the inserted hyperlink index.
- Add Aspose-style overloads for string range and row/column range anchors.
- Add `Hyperlink.LinkType`.
- Add `Hyperlink.Delete()`.
- Accept internal worksheet locations through the public `Address` API while
  still persisting SpreadsheetML `location` for internal links.
- Remove temporary public exposure of hyperlink sub-address internals.

### Page setup

- Align `PageSetup.LeftMargin` and related properties to centimeter semantics.
- Add inch-based variants `LeftMarginInch`, `RightMarginInch`, `TopMarginInch`,
  `BottomMarginInch`, `HeaderMarginInch`, and `FooterMarginInch`.
- Keep internal XLSX storage in inch units.

## Phase 3

### Workbook load/save

- Add `Workbook.Save(string, SaveFormat)`.
- Make `Workbook` implement `IDisposable`.
- Keep stream/file support unchanged.

## Tests

Every compatibility change must update:

- unit tests for the public member shape
- compatibility tests that use Aspose-style calling patterns
- OpenXML interoperability tests when serialization is affected

## Known Residual Gaps After This Pass

The following may still remain less than full Aspose.Cells parity:

- exact locale-specific display formatting
- full `LoadOptions` and `SaveOptions` object model compatibility
- larger `WorksheetCollection` and `Worksheet` API breadth beyond the supported view/protection/auto-filter v0.1 surface
- advanced hyperlink behaviors outside current XLSX scope
