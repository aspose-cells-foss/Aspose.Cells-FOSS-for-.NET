# Current Compatibility Gaps

## Purpose

This document lists the areas where the current `Aspose.Cells_FOSS`
implementation is still not fully consistent with:

- Microsoft Excel behavior
- Aspose.Cells for .NET public API or runtime behavior

This is a status document, not a delivery plan. The alignment plan remains in
`Spec/api_compatibility_alignment.md`.

## Scope

This document separates two kinds of differences:

1. supported v0.1 features that still have behavior gaps
2. product-scope gaps that are still outside v0.1

There are currently no intentional public bridge APIs in the supported surface.

## A. Known Gaps In Supported v0.1 Features

### 1. Cell display formatting is only partial

Status: Known gap

Current behavior:

- Cell.StringValue and Cell.DisplayStringValue now use separate paths in
  src/Aspose.Cells_FOSS/Cell.cs and src/Aspose.Cells_FOSS/DisplayTextFormatter.cs.
- StringValue returns a stable style-independent textual representation.
- DisplayStringValue now applies supported numeric display formats,
  including grouped decimals, percent, scientific notation, fractions,
  basic positive/negative/zero and conditional section selection, color
  section stripping, text sections, and a pragmatic subset of date/time tokens.
- Numeric and date formatting now use WorkbookSettings.Culture, with locale-directive overrides such as [$-409], [$-804], and special long-date/long-time directives like [$-F800] and [$-F400].
- Elapsed time formats and long-time directives now honor the active display culture for decimal separators and localized long-time patterns.
- The core style object model and styles.xml persistence now cover the supported font, fill, border, alignment, and number-format API surface; the remaining gap here is display-text rendering fidelity.

Why this is still not fully consistent:

- Excel display text is locale-sensitive.
- Aspose.Cells for .NET also exposes a richer display-text behavior through
  Cell.StringValue and Cell.DisplayStringValue.
- Complex Excel format codes are not fully rendered yet, especially:
  - richer Excel locale semantics beyond the supported workbook-culture and locale-directive subset
  - exact Excel calendar rendering edge cases and the remaining long tail of date/time tokens
  - color-aware rendering behavior beyond stripping non-display directives
  - accounting-style padding and other advanced Excel-specific layout tokens

Current code references:

- src/Aspose.Cells_FOSS/Cell.cs
- src/Aspose.Cells_FOSS/DisplayTextFormatter.cs

Official Aspose references:

- https://reference.aspose.com/cells/net/aspose.cells/cell/stringvalue/
- https://reference.aspose.com/cells/net/aspose.cells/cell/displaystringvalue/
- https://reference.aspose.com/cells/net/aspose.cells/cell/value/

### 2. StringValue versus DisplayStringValue still has simplified semantics

Status: Known gap

Current behavior:

- StringValue is now style-independent and DisplayStringValue is style-aware.
- The exact boundary between these two properties is still a pragmatic subset,
  not a full Aspose.Cells clone.

Why this is still not fully consistent:

- In Aspose.Cells for .NET, these two properties have richer edge-case behavior
  around formatted text, cached formula values, and locale-sensitive rendering.
- The current implementation intentionally favors deterministic invariant output
  over full Excel UI fidelity.

Current code references:

- src/Aspose.Cells_FOSS/Cell.cs
- src/Aspose.Cells_FOSS/DisplayTextFormatter.cs
### 3. Hyperlink behavior is functionally supported, but still narrower than Excel

Status: `Known gap`

Current behavior:

- External hyperlinks and internal worksheet-location hyperlinks are persisted
  to XLSX and loaded back.
- `Address`, `ScreenTip`, `TextToDisplay`, and internal `location` mapping are
  supported.

Why this is still not fully consistent:

- Excel applies built-in visual hyperlink styling behavior that is not modeled
  automatically here.
- Advanced hyperlink behaviors beyond current SpreadsheetML scope are not
  implemented.

Current code references:

- `src/Aspose.Cells_FOSS/Hyperlink.cs`
- `src/Aspose.Cells_FOSS/HyperlinkCollection.cs`
- `src/Aspose.Cells_FOSS/XlsxWorkbookHyperlinks.cs`

Official Aspose references:

- `https://reference.aspose.com/cells/net/aspose.cells/hyperlink/`
- `https://reference.aspose.com/cells/net/aspose.cells/hyperlinkcollection/add/`

### 4. Data validation support is limited to classic SpreadsheetML

Status: `Known gap`

Current behavior:

- `Worksheet.Validations`, `ValidationCollection`, and `Validation` support the
  classic SpreadsheetML `<dataValidations>` model.
- Core validation type, operator, formulas, prompts, error messages, and `sqref`
  persistence are supported.

Why this is still not fully consistent:

- Excel and Aspose.Cells for .NET cover a wider validation feature set.
- The current implementation does not cover x14 extensions, IME mode, prompt
  window placement metadata, or runtime enforcement of entered values.
- Validation behavior is preserved at the file model level, not as a full Excel
  input engine.

Current code references:

- `src/Aspose.Cells_FOSS/Validation.cs`
- `src/Aspose.Cells_FOSS/ValidationCollection.cs`
- `src/Aspose.Cells_FOSS/XlsxWorkbookValidations.cs`

Official Aspose references:

- `https://reference.aspose.com/cells/net/aspose.cells/validation/`
- `https://reference.aspose.com/cells/net/aspose.cells/validationcollection/`

### 5. Conditional formatting support is broader, but still not fully Excel-complete

Status: `Known gap`

Current behavior:

- `Worksheet.ConditionalFormattings`, `ConditionalFormattingCollection`,
  `FormatConditionCollection`, and `FormatCondition` now cover the main
  SpreadsheetML rule families implemented in the local feature set, including
  text rules, time-period rules, duplicate/unique values, top/bottom rules,
  above/below average, color scales, data bars, icon sets, classic `cellIs`,
  and `expression` rules.
- Differential styles through `dxfs` are loaded and saved for supported rules.

Why this is still not fully consistent:

- Excel and Aspose.Cells for .NET still expose a wider conditional formatting
  surface, especially x14 extensions and richer threshold/visual customization.
- The current implementation preserves file-level rule definitions only and does
  not evaluate or render conditional formatting.
- Data bar and icon-set persistence is limited to the main SpreadsheetML subset,
  not the full extension model used by newer Excel files.

Current code references:

- `src/Aspose.Cells_FOSS/ConditionalFormattingCollection.cs`
- `src/Aspose.Cells_FOSS/FormatConditionCollection.cs`
- `src/Aspose.Cells_FOSS/FormatCondition.cs`
- `src/Aspose.Cells_FOSS/XlsxWorkbookConditionalFormatting.cs`

Official Aspose references:

- `https://reference.aspose.com/cells/net/aspose.cells/worksheet/conditionalformattings/`
- `https://reference.aspose.com/cells/net/aspose.cells/formatconditioncollection/`
- `https://reference.aspose.com/cells/net/aspose.cells/formatcondition/`

### 6. `Worksheet`, `WorksheetCollection`, `WorkbookSettings`, and `PageSetup` only cover a subset of the Aspose surface

Status: `Known gap`

Current behavior:

- The project now exposes the Aspose-style members needed for the supported
  v0.1 scope, including `ActiveSheetIndex`, `ActiveSheetName`, `Date1904`,
  worksheet view members, worksheet protection baseline APIs, worksheet
  auto filter baseline APIs, hyperlink APIs, and margin properties with
  centimeter and inch variants.
Why this is still not fully consistent:

- Aspose.Cells for .NET exposes a much broader object model in these classes.
- The current implementation intentionally covers only the supported v0.1
  feature set and does not yet replicate the wider member surface.

Current code references:

- `src/Aspose.Cells_FOSS/Worksheet.cs`
- `src/Aspose.Cells_FOSS/WorksheetProtection.cs`
- `src/Aspose.Cells_FOSS/AutoFilter.cs`
- `src/Aspose.Cells_FOSS/AutoFilterCriteria.cs`
- `src/Aspose.Cells_FOSS/WorksheetCollection.cs`
- `src/Aspose.Cells_FOSS/WorkbookSettings.cs`
- `src/Aspose.Cells_FOSS/DefinedNameCollection.cs`
- `src/Aspose.Cells_FOSS/DefinedName.cs`
- `src/Aspose.Cells_FOSS/PageSetup.cs`
- `src/Aspose.Cells_FOSS/XlsxWorkbookWorksheetViews.cs`
- `src/Aspose.Cells_FOSS/XlsxWorkbookWorksheetProtection.cs`

Official Aspose references:

- `https://reference.aspose.com/cells/net/aspose.cells/worksheetcollection/`
- `https://reference.aspose.com/cells/net/aspose.cells/workbooksettings/`
- `https://reference.aspose.com/cells/net/aspose.cells/worksheet/autofilter/`
- `https://reference.aspose.com/cells/net/aspose.cells/pagesetup/`

### 7. Save and load option models are still FOSS-specific

Status: `Known gap`

Current behavior:

- `LoadOptions` and `SaveOptions` are project-specific types oriented around the
  current XLSX implementation.
- `SaveFormat` and `LoadFormat` currently only expose XLSX-oriented values.

Why this is still not fully consistent:

- Aspose.Cells for .NET has a much larger and more format-rich options system.
- Current option names and behaviors are compatible only where needed for the
  supported v0.1 scope.

Current code references:

- `src/Aspose.Cells_FOSS/LoadOptions.cs`
- `src/Aspose.Cells_FOSS/SaveOptions.cs`
- `src/Aspose.Cells_FOSS/LoadFormat.cs`
- `src/Aspose.Cells_FOSS/SaveFormat.cs`

Official Aspose references:

- `https://reference.aspose.com/cells/net/aspose.cells/loadoptions/`
- `https://reference.aspose.com/cells/net/aspose.cells/saveoptions/`
- `https://reference.aspose.com/cells/net/aspose.cells/workbook/save/`

### 8. Some public typing still reflects FOSS-specific simplification

Status: `Known gap`

Current behavior:

- `Cell.Type` returns the local `CellValueType` enum.

Why this is still not fully consistent:

- This is a practical API for the current v0.1 implementation, but it should
  not be treated as proof that the full Aspose.Cells cell typing model has been
  replicated.
- It represents a reduced surface aligned to the current product scope.

Current code references:

- `src/Aspose.Cells_FOSS/Cell.cs`
- `src/Aspose.Cells_FOSS/CellValueType.cs`

## B. Product-Scope Differences From Excel And Aspose.Cells

Status: `Out of scope in v0.1`

The following remain outside the current product scope and therefore are still
major differences from both Excel and Aspose.Cells for .NET:

- formula calculation engine
- charts
- images and drawings
- pivot tables
- tables
- comments
- macros and VBA
- encryption
- streaming writer
- non-XLSX file-format coverage
- x14 conditional formatting extensions and richer threshold/visual customization beyond the main SpreadsheetML subset

Scope references:

- `Spec/product_scope.md`

## C. Summary

The current codebase is closer to Aspose.Cells for .NET than earlier revisions,
but it is still not fully identical in the following practical areas:

- display-text formatting fidelity
- full Excel-level conditional formatting fidelity
- full public API breadth
- option-object compatibility
- broader Excel feature coverage outside v0.1













