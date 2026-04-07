# Implementation Steps

## Compatibility Alignment

Before expanding feature breadth, the supported public API surface should be
aligned to Aspose.Cells for .NET for the features already implemented in v0.1.

Preferred order:

1. `Cell`
2. `Worksheet` and `WorkbookSettings`
3. `WorksheetCollection`
4. `HyperlinkCollection` and `Hyperlink`
5. `PageSetup`
6. `Workbook` convenience and lifetime members

## 1. Cell Data APIs and Excel Cell Data Import/Export

### Goal
Implement public cell data APIs and XLSX import/export for core scalar cell values:
- `bool`
- `int`
- `decimal`
- `double`
- `string`
- `DateTime`
- blank / null-like cell state
- formula text persistence

### Step 1. Define public API surface
- Finalize `Cell` read/write APIs for scalar values.
- Add overloads for `PutValue(bool)`, `PutValue(int)`, `PutValue(decimal)`, `PutValue(double)`, `PutValue(string)`, and `PutValue(DateTime)`.
- Define behavior for `Value`, `StringValue`, `Formula`, and `Type`.
- Define null/blank handling rules and exception behavior.

### Step 2. Normalize internal cell value model
- Introduce a stable internal value representation for scalar types.
- Separate logical cell kind from XML storage form.
- Keep formula text and cached value handling independent.
- Ensure cell storage supports sparse access by row/column and A1 address.

### Step 3. Implement address and index conversion
- Complete A1 parsing and formatting.
- Validate zero-based public row/column indexers.
- Guarantee conversion correctness for columns beyond `Z`.
- Add guard rails for invalid references and negative indexes.

### Step 4. Implement workbook date system support
- Route all `DateTime` conversion through a single serial conversion service.
- Support both 1900 and 1904 date systems from workbook settings.
- Define import and export behavior for date-like numeric cells.
- Preserve numeric storage with date formatting semantics separated from value semantics.

### Step 5. Implement cell import from worksheet XML
- Read `<c>`, `<v>`, `<f>`, and inline/shared string content.
- Map XML cell types to internal value kinds.
- Load booleans, numbers, strings, and date-related numeric values.
- Recover safely from missing or inconsistent optional attributes where allowed by policy.

### Step 6. Implement shared string integration
- Build shared string table loading and lookup.
- Resolve shared string indexes during import.
- Allocate stable shared string indexes during save.
- Support inline strings as alternate input/output mode.

### Step 7. Implement cell export to worksheet XML
- Serialize scalar values using correct SpreadsheetML forms.
- Emit formula cells without calculation engine behavior.
- Write booleans, numbers, strings, and dates deterministically.
- Omit blank cells unless required by style, formula, or merge semantics.

### Step 8. Add recovery and diagnostics
- Detect broken shared string references, invalid cached values, and malformed references.
- Record warnings in `LoadDiagnostics`.
- Apply repair rules for recoverable cases defined in the specs.
- Mark lossy recovery cases explicitly.

### Step 9. Add tests
- Unit tests for type mapping, A1 parsing, date conversion, and `StringValue`.
- Golden tests for mixed-type worksheets and round-trip save/load.
- Malformed tests for broken shared strings, invalid references, and unsorted cells.
- Compatibility tests for public API behavior and file/stream parity.

### Step 10. Deliverable checkpoint
- Public cell data APIs complete.
- XLSX import/export works for the target scalar types.
- Date system behavior is correct for both 1900 and 1904 modes.
- Diagnostics exist for recoverable and lossy input cases.

## 2. Style APIs and Full Style Import/Export

### Goal
Implement public style APIs and XLSX import/export for style settings including:
- Font
- Borders
- Alignments
- Number formats

### Step 1. Finalize public style object model
- Complete the `Style`, `Font`, `Borders`, and `Border` APIs.
- Confirm naming and behavior compatibility with Aspose-style usage.
- Define clone/copy-on-write behavior for style mutation.
- Define default style semantics and style identity rules.

### Step 2. Finalize internal style value model
- Separate public style objects from internal immutable or value-based style records.
- Represent font, border, alignment, protection, fill, and number format values independently.
- Add a style repository for deduplication and stable style indexing.
- Ensure default style index `0` always exists.

### Step 3. Implement font import/export
- Read font settings from `styles.xml`.
- Map name, size, bold, italic, underline, and color fields.
- Serialize deterministic font collections on save.
- Deduplicate equivalent font definitions.

### Step 4. Implement border import/export
- Read left, right, top, and bottom border settings.
- Map line style and color for each border side.
- Serialize stable border pools and cell XF references.
- Preserve default border behavior when no explicit border is set.

### Step 5. Implement alignment import/export
- Read horizontal alignment, vertical alignment, and wrap text settings.
- Apply alignment values through public `Style` APIs.
- Emit alignment nodes only when required.
- Validate supported values and fallback behavior for unsupported cases.

### Step 6. Implement number format import/export
- Load built-in and custom number format definitions.
- Map format index and custom format string through `Style`.
- Preserve deterministic format ID allocation on save.
- Use number formats to support date display semantics without changing underlying numeric storage.

### Step 7. Implement style binding between cells and stylesheet
- Attach style values to cells during import.
- Resolve cell XF indexes safely, including default fallback behavior.
- Assign stable style indexes during save.
- Omit explicit style index when the cell uses the default style.

### Step 8. Implement style recovery and diagnostics
- Handle missing stylesheet with default-style-only workbooks.
- Map invalid style indexes to style `0` with warnings.
- Record unsupported advanced style features as lossy recoverable cases.
- Preserve future extension data where configured and safe.

### Step 9. Add tests
- Unit tests for style equality, deduplication, and copy-on-write mutation behavior.
- Golden tests for fonts, borders, alignments, and number formats.
- Malformed tests for invalid style indexes and inconsistent stylesheet counts.
- Compatibility tests for style getter/setter behavior on cells.

### Step 10. Deliverable checkpoint
- Public style APIs are usable and stable.
- Import/export works for fonts, borders, alignments, and number formats.
- Style indexing is deterministic.
- Recovery and diagnostics exist for invalid or incomplete style data.

## 3. Worksheet Options and Settings APIs and Import/Export

### Goal
Implement public worksheet options/settings APIs and XLSX import/export for worksheet-level configuration including:
- worksheet name and visibility
- row and column settings
- default row / column behavior
- merged cells / ranges
- data validation
- worksheet view and pane basics
- sheet protection baseline settings

### Step 1. Finalize public worksheet settings API surface
- Define worksheet-level APIs for visibility and basic sheet options.
- Add row and column settings APIs needed by the spec.
- Define merged range APIs and worksheet option access patterns.
- Confirm behavior for default values and mutation semantics.

### Step 2. Finalize internal worksheet options model
- Extend the worksheet model to store row, column, merge, view, and protection settings.
- Separate sheet-level metadata from cell-level data.
- Define sparse storage rules for row and column records.
- Ensure defaults are stable when options are omitted in source XML.

### Step 3. Implement worksheet metadata import/export
- Read and write worksheet name, sheet state, and workbook sheet bindings.
- Import and export visible, hidden, and very hidden states.
- Preserve deterministic sheet ordering.
- Validate duplicate or invalid worksheet names.

### Step 4. Implement row and column settings import/export
- Read row height, hidden state, and row style references.
- Read column width, hidden state, and column style references.
- Serialize row and column records only when required.
- Preserve default sizing behavior when explicit values are absent.

### Step 5. Implement merged cell import/export
- Read merged cell ranges from worksheet XML.
- Validate merged range references and overlap behavior.
- Serialize merged ranges deterministically.
- Ensure merge settings are preserved independently of cell value persistence.

### Step 6. Implement worksheet view and pane basics
- Read basic sheet view state, active cell, selected ranges, and freeze/split pane metadata as scoped by the spec.
- Map supported options to public APIs.
- Serialize worksheet views deterministically.
- Ignore or diagnose unsupported advanced view features.

### Step 7. Implement worksheet protection baseline support
- Read and write worksheet protection flags required by the spec.
- Preserve supported protection options through public APIs.
- Record unsupported protection details as diagnostics when needed.
- Keep sheet protection behavior independent from workbook protection.

### Step 8. Add recovery and diagnostics
- Handle malformed merge ranges, invalid row/column spans, and broken worksheet settings safely.
- Record warnings for recoverable worksheet option issues.
- Define fallback behavior for missing optional worksheet settings nodes.
- Mark lossy worksheet-option recovery cases explicitly.

### Step 9. Add tests
- Unit tests for visibility, merge APIs, data validation APIs, and row/column option mutation.
- Golden tests for worksheet settings and data validations round-trip save/load.
- Malformed tests for invalid merge ranges, invalid validation sqref/type, and broken row/column metadata.
- Compatibility tests for worksheet option APIs, validation APIs, and deterministic output behavior.

### Step 10. Deliverable checkpoint
- Worksheet options/settings APIs are usable and stable.
- Import/export works for the targeted worksheet metadata and options.
- Merge, row, and column settings round-trip correctly.
- Recovery and diagnostics exist for malformed worksheet settings.

## 4. Page Setup APIs and Import/Export

### Goal
Implement public page setup APIs and XLSX import/export for print and page-layout settings including:
- paper size and orientation
- margins
- scaling / fit-to-page options
- print area and print titles
- header / footer basics
- page breaks and print options baseline

### Step 1. Finalize public page setup API surface
- Define a `PageSetup` API model attached to each worksheet.
- Add APIs for margins, orientation, paper size, scaling, and fit-to-page.
- Define print area, repeating rows/columns, and header/footer access patterns.
- Confirm defaults and mutation semantics for page setup state.

### Step 2. Finalize internal page setup model
- Add internal models for page setup, print options, margins, and header/footer data.
- Separate worksheet page settings from worksheet view and cell data.
- Define stable defaults for absent print-related XML.
- Ensure workbook-defined names can back print area and print titles.

### Step 3. Implement page margins import/export
- Read left, right, top, bottom, header, and footer margins.
- Map margin values through public APIs with deterministic numeric handling.
- Serialize page margins only when required.
- Preserve default margin behavior when not explicitly set.

### Step 4. Implement page setup core import/export
- Read paper size, orientation, first page number, scale, fit-to-width, and fit-to-height.
- Map supported page setup attributes to public APIs.
- Serialize page setup attributes deterministically.
- Validate unsupported or malformed values with fallback behavior.

### Step 5. Implement print area and print titles import/export
- Read workbook-defined names for print area and repeating rows/columns.
- Resolve those names to worksheet-scoped page setup APIs.
- Serialize print area and print title definitions back to workbook metadata.
- Preserve sheet-scoped name binding and deterministic naming rules.

### Step 6. Implement header/footer and print options baseline
- Read and write supported header/footer strings.
- Import and export print gridlines, headings, centering, and other baseline print options from the spec.
- Preserve deterministic XML emission for supported options.
- Record unsupported advanced header/footer constructs as diagnostics where necessary.

### Step 7. Implement page break baseline support
- Read horizontal and vertical page break metadata in scope for the spec.
- Expose supported page break APIs or internal preservation behavior.
- Serialize page breaks deterministically.
- Validate invalid break references safely.

### Step 8. Add recovery and diagnostics
- Handle malformed defined names, invalid margin values, and broken page setup metadata safely.
- Record warnings for recoverable page setup issues.
- Apply fallback behavior for missing or inconsistent print settings.
- Mark lossy recovery cases explicitly.

### Step 9. Add tests
- Unit tests for page setup API mutation and default values.
- Golden tests for margins, orientation, scaling, print area, and print titles round-trip behavior.
- Malformed tests for broken defined names and invalid page setup attributes.
- Compatibility tests for file/stream parity and deterministic serialization of page settings.

### Step 10. Deliverable checkpoint
- Public page setup APIs are usable and stable.
- Import/export works for targeted page and print settings.
- Workbook-defined print metadata round-trips correctly.
- Recovery and diagnostics exist for malformed page setup content.

## 5. Cross-Feature OpenXML Interoperability Validation

### Goal
Add a single OpenXML-based verification scenario that generates one workbook combining every currently supported v0.1 feature and checks the emitted package against the original Aspose.Cells_FOSS configuration.

### Required coverage
- workbook settings such as date system and worksheet visibility
- cell values including string, number, boolean, DateTime, and formula cached value persistence
- core style persistence for font, fill, borders, alignment, number format, and protection flags
- worksheet settings including rows, columns, merges, dimension, hyperlinks, data validations, and conditional formatting
- page setup including margins, print titles, print area, headers/footers, and page breaks

### Deliverable checkpoint
- Aspose.Cells_FOSS can generate a workbook containing all currently supported feature settings
- Open XML SDK inspection matches the original Aspose.Cells_FOSS settings for that workbook
- the interoperability test runs as part of the repository test suite



