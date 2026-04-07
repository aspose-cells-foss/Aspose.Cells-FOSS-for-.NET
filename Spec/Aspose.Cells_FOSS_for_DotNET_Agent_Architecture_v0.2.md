# Aspose.Cells FOSS for .NET â€?Agent Architecture & Engineering Plan v0.2

## 1. Project Positioning

**Project name:** Aspose.Cells FOSS for .NET  
**Namespace:** `Aspose.Cells_FOSS`

This project aims to build a **commercial-grade**, **fully self-hosted**, **pure managed .NET** library for manipulating Excel XLSX files according to **ECMA-376 / SpreadsheetML / OPC**, while keeping the **public object model highly compatible with Aspose.Cells for .NET**.

### Core goals
- Commercial product quality
- Fully self-developed low-level implementation
- Public API close to Aspose.Cells for .NET
- XLSX-focused for v0.1
- File and stream loading/saving from day one
- Recovery-friendly loading for damaged but repairable XLSX files
- 1900 and 1904 date systems supported from the beginning
- Full `Style` object API surface defined in v0.1

---

## 2. Product Requirements Already Confirmed

### Product and architecture
- Product target: **commercial-grade**
- Low-level implementation: **fully self-developed**
- API style: **as close as possible to Aspose.Cells for .NET**, with room for additions
- Target frameworks: **`netstandard2.0` and `net8.0`**
- Streaming writer: **not required in v0.1**

### Functional scope
- v0.1 focuses on **core XLSX reading/writing**
- No charts/images/conditional formatting in v0.1
- File path and `Stream` input/output both required
- Both **1900** and **1904** date systems required
- `Style` object API should be **fully defined** in v0.1

### Robustness and compatibility
- Must support loading **damaged but recoverable XLSX files**
- Exception types/behaviors should be **as compatible with Aspose.Cells as practical**

---

## 3. Recommended v0.1 Scope

### Included
#### Workbook / worksheet
- Create workbook
- Open workbook from file
- Open workbook from stream
- Save workbook to file
- Save workbook to stream
- Worksheet collection
- Add / remove / rename sheets
- Basic sheet visibility
- Workbook settings
- 1900 / 1904 date system

#### Cells
- Access by A1 notation
- Access by row/column
- Read/write cell values
- Strings
- Numbers
- Booleans
- Date / DateTime
- Formulas (persistence only, no calculation engine)
- Blank cells
- Shared strings
- Inline strings

#### Structural worksheet features
- Row metadata
- Column metadata
- Row height
- Column width
- Hidden rows
- Hidden columns
- Merge cells
- Sheet dimension

#### Styles
- Full public `Style` API surface planned
- v0.1 persisted core:
  - Font
  - Fill
  - Border
  - Number format
  - Horizontal alignment
  - Vertical alignment
  - Wrap text
  - Locked / hidden

#### Load robustness
- Strict mode
- Relaxed / repair-friendly mode
- Load diagnostics
- Recovery and warning reporting

### Excluded from v0.1
- Charts
- Images / drawings
- Conditional formatting
- Data validation
- Pivot tables
- Tables
- Comments / notes
- VBA / macros
- Encryption
- Formula calculation engine
- Streaming writer

---

## 4. Core Design Principles

### 4.1 Object model first
The public API must be designed from an **Aspose.Cells-style object model**, not from raw XML structures.

Target usage style:

```csharp
var workbook = new Workbook();
var sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("Hello");
sheet.Cells["B1"].PutValue(123);
workbook.Save("demo.xlsx");
```

Not like this:

```csharp
worksheetPart.SheetData.Rows[0].Cells[0].Value = ...
```

### 4.2 Spec-driven development
The engineering process should follow this chain:

**ECMA-376 / SpreadsheetML / OPC**  
â†?structured spec files  
â†?architecture and API design  
â†?code generation / implementation  
â†?validation and regression testing

### 4.3 No production dependency on Open XML SDK
Production code should not depend on Open XML SDK.

It may still be used in:
- test tooling
- interoperability verification
- regression validation
- structure comparison

### 4.4 Separate compatibility from implementation
The project should behave similarly to Aspose.Cells at the public API level, while using a modern internal architecture that is easier to maintain, validate, and evolve.

---

## 5. Recommended Agent System

To make the project scalable, use an agent pipeline rather than a single â€œcode writingâ€?agent.

### Agent 1 â€?Scope Agent
Responsible for:
- defining version boundaries
- feature prioritization
- roadmap planning
- backlog management

Outputs:
- `roadmap.md`
- `feature_backlog.yaml`

### Agent 2 â€?Aspose Compatibility Agent
Responsible for:
- mapping required Aspose.Cells public APIs
- documenting desired behavior compatibility
- documenting exception compatibility

Outputs:
- `aspose_api_matrix.yaml`
- `behavior_compatibility.md`
- `exception_compatibility.md`

### Agent 3 â€?Spec Agent
Responsible for:
- decomposing ECMA-376 / OPC / SpreadsheetML rules
- converting prose specifications into machine-readable structured specs
- defining XML mappings, constraints, defaults, and validation rules

Outputs:
- `specs/opc/*.yaml`
- `specs/spreadsheetml/*.yaml`

### Agent 4 â€?Recovery / Robustness Agent
Responsible for:
- defining recoverable parsing rules
- damaged file handling strategy
- diagnostics and repair action models

Outputs:
- `recovery_policy.yaml`
- `load_diagnostics_spec.yaml`
- `malformed_cases_catalog.yaml`

### Agent 5 â€?Architecture Agent
Responsible for:
- defining layered architecture
- mapping public API to internal domain models
- defining loading and saving pipelines
- defining style system strategy and storage models

Outputs:
- `architecture.md`
- `domain_model.yaml`
- `module_graph.md`

### Agent 6 â€?Code Generator Agent
Responsible for:
- generating solution skeleton
- generating class skeletons
- generating parser/writer/validator skeletons
- generating test skeletons

Outputs:
- `src/`
- `tests/`

### Agent 7 â€?Verification Agent
Responsible for:
- golden tests
- round-trip tests
- malformed file tests
- compatibility tests
- conformance reports

Outputs:
- `conformance_report.md`
- `compatibility_report.md`

### Agent 8 â€?Regression Agent
Responsible for:
- checking public API compatibility drift
- checking XML output drift
- checking style/shared-string index stability
- checking exception behavior regressions

Outputs:
- `api_regression_report.md`
- `binary_compatibility_report.md`

---

## 6. Recommended Architecture

### 6.1 Public object model layer
Exposed to library users.

Representative classes:
- `Workbook`
- `WorksheetCollection`
- `Worksheet`
- `Cells`
- `Cell`
- `Style`
- `Font`
- `Borders`
- `Border`
- `Range`
- `LoadOptions`
- `SaveOptions`

### 6.2 Facade / behavior layer
Bridges public API calls into internal domain services.

Representative types:
- `WorkbookFacade`
- `WorksheetFacade`
- `CellsFacade`
- `StyleFacade`

### 6.3 Internal domain model layer
The real internal source of truth.

Representative types:
- `WorkbookModel`
- `WorksheetModel`
- `CellStore`
- `CellRecord`
- `WorkbookSettingsModel`
- `StyleRepository`
- `StyleValue`
- `SharedStringTableModel`
- `PackageModel`

### 6.4 Feature service layer
Responsible for business logic and normalization.

Representative services:
- `CellValueService`
- `DateSerialConverter`
- `SharedStringService`
- `FormulaService`
- `StyleService`
- `MergeCellService`
- `RowColumnService`
- `WorksheetDimensionService`

### 6.5 Serialization / parsing layer
Responsible for XML â†?model mapping.

Representative components:
- `WorkbookXmlReader`
- `WorkbookXmlWriter`
- `WorksheetXmlReader`
- `WorksheetXmlWriter`
- `StylesXmlReader`
- `StylesXmlWriter`
- `SharedStringsXmlReader`
- `SharedStringsXmlWriter`
- `RelationshipsXmlReader`
- `RelationshipsXmlWriter`
- `ContentTypesXmlReader`
- `ContentTypesXmlWriter`

### 6.6 OPC packaging layer
Responsible for ZIP package manipulation and package relationship structure.

Representative components:
- `ZipPackageReader`
- `ZipPackageWriter`
- `PartRegistry`
- `PartUriAllocator`
- `RelationshipManager`
- `ContentTypeManager`

### 6.7 Validation layer
Responsible for correctness checking and conformance verification.

Representative validators:
- `StructuralValidator`
- `SemanticValidator`
- `ReferenceValidator`
- `ConformanceValidator`

### 6.8 Recovery model layer
Required because damaged-but-recoverable loading is a first-class requirement.

Representative types:
- `LoadDiagnostics`
- `LoadIssue`
- `RepairAction`
- `PackageIssue`
- `XmlIssue`
- `SemanticIssue`

---

## 7. Public API Direction

The public API should intentionally feel like Aspose.Cells.

### Workbook
```csharp
public class Workbook
{
    public Workbook();
    public Workbook(string fileName);
    public Workbook(Stream stream);
    public Workbook(string fileName, LoadOptions options);
    public Workbook(Stream stream, LoadOptions options);

    public WorksheetCollection Worksheets { get; }
    public WorkbookSettings Settings { get; }
    public LoadDiagnostics LoadDiagnostics { get; }

    public void Save(string fileName);
    public void Save(string fileName, SaveOptions options);
    public void Save(Stream stream, SaveFormat format);
    public void Save(Stream stream, SaveOptions options);
}
```

### Worksheet collection
```csharp
public class WorksheetCollection
{
    public Worksheet this[int index] { get; }
    public Worksheet this[string name] { get; }

    public int Add(string sheetName);
    public void RemoveAt(int index);
}
```

### Worksheet
```csharp
public class Worksheet
{
    public string Name { get; set; }
    public Cells Cells { get; }
}
```

### Cells
```csharp
public class Cells
{
    public Cell this[string cellName] { get; }
    public Cell this[int row, int column] { get; }

    public void Merge(int firstRow, int firstColumn, int totalRows, int totalColumns);
}
```

### Cell
```csharp
public class Cell
{
    public object? Value { get; }
    public string StringValue { get; }
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

### Style
The public `Style` object should have a broad API surface from v0.1, even if some members are persisted in later phases.

---

## 8. Load / Save Design

### 8.1 LoadOptions
Recommended design:

```csharp
public class LoadOptions
{
    public LoadFormat LoadFormat { get; set; } = LoadFormat.Auto;
    public bool StrictMode { get; set; } = false;
    public bool TryRepairPackage { get; set; } = true;
    public bool TryRepairXml { get; set; } = true;
    public bool PreserveUnsupportedParts { get; set; } = true;
    public IWarningCallback? WarningCallback { get; set; }
}
```

### 8.2 SaveOptions
Recommended design:

```csharp
public class SaveOptions
{
    public SaveFormat SaveFormat { get; set; } = SaveFormat.Xlsx;
    public bool UseSharedStrings { get; set; } = true;
    public bool ValidateBeforeSave { get; set; } = true;
    public bool CompactStyles { get; set; } = true;
    public bool PreserveRecoveryMetadata { get; set; } = false;
}
```

### 8.3 LoadDiagnostics
Recommended design:

```csharp
public sealed class LoadDiagnostics
{
    public IReadOnlyList<LoadIssue> Issues { get; }
    public bool HasRepairs { get; }
    public bool HasDataLossRisk { get; }
}
```

Usage example:

```csharp
var workbook = new Workbook(stream, new LoadOptions
{
    TryRepairPackage = true,
    TryRepairXml = true
});

var diagnostics = workbook.LoadDiagnostics;
```

---

## 9. Recovery-Friendly Loading Strategy

Because damaged-but-recoverable files must be supported, recovery cannot be handled only with exceptions.

### 9.1 Error categories

#### Fatal
Cannot continue loading.
Examples:
- file is not a ZIP package
- package cannot be enumerated
- workbook part is missing and cannot be inferred

Action:
- throw exception

#### Recoverable
Can be repaired and loading can continue.
Examples:
- missing relationship that can be inferred
- invalid or missing dimension that can be recalculated
- rows or cells out of order
- shared string count inconsistency when data is still readable
- ignorable extra nodes

Action:
- record issue
- apply repair
- continue loading

#### Lossy recoverable
Loading can continue, but some information may be lost or downgraded.
Examples:
- style index out of range
- unsupported extension content
- unsupported feature parts preserved but not interpreted
- unsupported drawing parts in v0.1

Action:
- record warning
- preserve data when possible
- mark data loss risk

### 9.2 Recovery layers

#### Package-level repair
- infer missing content type overrides when possible
- infer missing relationships using well-known part paths
- downgrade optional missing parts

#### XML-level repair
- reorder rows and cells
- ignore invalid attributes when safe
- normalize element ordering
- recalculate dimensions
- accept duplicated nodes with deterministic resolution and diagnostics

#### Semantic-level repair
- fallback to default style for invalid style index
- fallback behavior for broken shared string references
- preserve formula text while dropping invalid cached values

---

## 10. Exception Design

To stay compatible with Aspose-like external behavior while keeping the internals maintainable, use two exception layers.

### 10.1 Public exception layer
Examples:
- `CellsException`
- `InvalidFileFormatException`
- `UnsupportedFeatureException`
- `WorkbookLoadException`
- `WorkbookSaveException`
- `FormulaException`
- `StyleException`

### 10.2 Internal exception layer
Examples:
- `PackageStructureException`
- `MissingPartException`
- `RelationshipResolutionException`
- `XmlParsingException`
- `StyleIndexOutOfRangeException`
- `SharedStringCorruptionException`

Internal exceptions should be mapped to public exceptions at the API boundary.

---

## 11. Style System Strategy

Because the full `Style` API should exist from v0.1, style design must be future-proof.

### 11.1 Key principle
**Public style objects should not be the internal storage format.**

Recommended model:
- public `Style` object = facade / handle
- internal `StyleValue` = immutable value object
- `StyleRepository` = deduplication and pooling layer
- style IDs assigned during serialization

### 11.2 Why this matters
This makes it easier to:
- deduplicate styles
- keep stable style indexes
- compare styles efficiently
- reduce file size
- support copy-on-write behavior

### 11.3 Style implementation levels

#### API complete
Expose the broad `Style` API shape from v0.1.

#### Persisted core
Actually read/write these first:
- font name / size / bold / italic / underline / color
- fill pattern / foreground / background
- border styles / colors
- number format
- horizontal alignment
- vertical alignment
- wrap text
- locked / hidden

#### Deferred behavior
Some public properties may initially be:
- retained in memory only
- ignored during save
- documented as not yet persisted

---

## 12. Date System Design

Both **1900** and **1904** date systems must be supported from the beginning.

### Recommended model
```csharp
public sealed class WorkbookSettingsModel
{
    public DateSystem DateSystem { get; set; } = DateSystem.Windows1900;
}
```

### Required coverage
- serial â†?`DateTime`
- `DateTime` â†?serial
- `workbookPr` date1904 flag read/write
- formula cached value consistency
- integration with number formatting for dates

### Important rule
All date conversions should go through a single service:

- `DateSerialConverter`

Do not spread date conversion logic throughout the codebase.

---

## 13. Internal Data Structures

### 13.1 Cell storage
Do not use a giant 2D array.

Recommended direction:
- sparse worksheet store
- row-oriented storage
- sorted cells within each row

Example shape:
- `Dictionary<int, RowModel>`
- `SortedDictionary<int, CellRecord>`

This is better for:
- memory efficiency
- sparse worksheets
- serialization order
- future streaming support

### 13.2 Shared strings
Recommended internal model:
- `Dictionary<string, int>` for value-to-index
- `List<SharedStringEntry>` for index-to-value

Shared strings should be enabled by default, with an option to use inline strings.

### 13.3 Style storage
Use pooled repositories:
- font pool
- fill pool
- border pool
- number format pool
- cell format pool

---

## 14. Solution Structure

```text
Aspose.Cells_FOSS/
  docs/
    product_scope.md
    roadmap.md

    architecture/
      architecture.md
      module_graph.md
      loading_pipeline.md
      saving_pipeline.md
      recovery_design.md
      style_system.md

    api/
      public_api.md
      aspose_compatibility.md
      exception_compatibility.md

    specs/
      opc/
        package.yaml
        content_types.yaml
        relationships.yaml

      spreadsheetml/
        workbook.yaml
        worksheet.yaml
        cells.yaml
        rows.yaml
        columns.yaml
        merges.yaml
        styles.yaml
        shared_strings.yaml
        formulas.yaml
        dates.yaml

      recovery/
        recovery_policy.yaml
        malformed_cases_catalog.yaml

    adr/
      ADR-001-public-api-compatible-with-aspose.md
      ADR-002-self-hosted-opc-and-xml-stack.md
      ADR-003-sparse-cell-store.md
      ADR-004-style-pool-copy-on-write.md
      ADR-005-recovery-friendly-loader.md
      ADR-006-dual-target-netstandard20-net80.md

  agents/
    prompts/
      scope_agent.md
      compatibility_agent.md
      spec_agent.md
      recovery_agent.md
      architecture_agent.md
      codegen_agent.md
      verification_agent.md

    schemas/
      feature_spec.schema.json
      api_model.schema.json
      recovery_rule.schema.json
      test_matrix.schema.json

    workflows/
      feature_intake.md
      spec_compile.md
      architecture_mapping.md
      codegen.md
      verify.md

  src/
    Aspose.Cells_FOSS/
      Aspose.Cells_FOSS.csproj
      Cell.cs
      LoadSave.cs
      Style.cs
      Workbook.cs
      WorkbookSettings.cs
      Worksheet.cs
      XlsxWorkbookSerializer.cs
      XlsxWorkbookSerializer.Helpers.cs

      Core/
        DateSerialConverter.cs
        Models.cs

      Packaging/
        PackagingSkeleton.cs

      Xml/
        SpreadsheetMlSkeleton.cs

      Validation/
        ValidationSkeleton.cs

  tests/
    Aspose.Cells_FOSS.UnitTests/
    Aspose.Cells_FOSS.GoldenTests/
    Aspose.Cells_FOSS.MalformedTests/
    Aspose.Cells_FOSS.CompatibilityTests/

  samples/
    Aspose.Cells_FOSS.Samples.Basic/
    Aspose.Cells_FOSS.Samples.Loading/
    Aspose.Cells_FOSS.Samples.Styles/
```

---

## 15. Project Breakdown

### `Aspose.Cells_FOSS`
Single library project and single runtime assembly.

Responsibilities:
- public API surface
- internal workbook/worksheet/cell model
- XLSX packaging and SpreadsheetML read/write logic
- validation and diagnostics
- save/load options, exceptions, and settings

### Internal source layout inside `Aspose.Cells_FOSS`
The source is organized by folders inside the single project rather than separate library projects.

Responsibilities by folder:
- `Core/`: workbook models, style values, shared strings, date conversion, diagnostics
- `Packaging/`: packaging conventions and package-level placeholders
- `Xml/`: SpreadsheetML XML mapping placeholders
- `Validation/`: validation placeholders and future conformance logic

---

## 16. Multi-Targeting Strategy

The main library should target:

```xml
<TargetFrameworks>netstandard2.0;net8.0</TargetFrameworks>
```

### Strategy
- Keep public API consistent across both targets
- Use conditional optimization paths internally for `net8.0`
- Keep baseline implementation portable for `netstandard2.0`

### Expected benefits
- wide compatibility through `.NET Standard 2.0`
- better runtime performance on `.NET 8`
- future-proof modernization path without sacrificing reach

---

## 17. Spec System Recommendations

The spec files are one of the most valuable long-term assets for this project.

### Recommended spec folders
- `docs/specs/opc/`
- `docs/specs/spreadsheetml/`
- `docs/specs/recovery/`
- `docs/api/`
- `docs/adr/`

### Example feature spec
```yaml
feature: worksheet.basic.cells
version: 1
scope:
  include:
    - text
    - number
    - bool
    - datetime
    - formula
public_api:
  classes:
    - Cells
    - Cell
  methods:
    - Cells.this[string]
    - Cells.this[int,int]
    - Cell.PutValue
    - Cell.GetStyle
    - Cell.SetStyle
internal_model:
  entities:
    - WorksheetModel
    - CellRecord
    - CellValue
xml_mapping:
  part: /xl/worksheets/sheet{n}.xml
  elements:
    - worksheet
    - sheetData
    - row
    - c
    - v
    - f
constraints:
  - rows ordered ascending
  - cells ordered ascending within row
  - c@r valid A1 reference
recovery:
  - reorder rows when possible
  - reorder cells when possible
  - ignore unknown elements in supported safe zones
tests:
  - write_string
  - write_number
  - write_bool
  - write_datetime
  - write_formula
  - roundtrip
  - malformed_row_order
```

### Example recovery rule
```yaml
rule_id: worksheet.rows.unsorted
severity: recoverable
detect:
  condition: row indices not ascending
repair:
  action: reorder rows by row index
diagnostics:
  warning_code: ACF0012
  message: Worksheet rows were out of order and have been normalized.
```

### Example exception mapping
```yaml
internal_exception: MissingWorkbookPartException
public_exception: WorkbookLoadException
error_code: ACF1001
recoverable: false
```

---

## 18. Testing Strategy

A commercial-grade XLSX library needs a layered testing system.

### 18.1 Unit tests
Test pure logic:
- A1 address parsing
- row/column conversions
- date serial conversion
- style deduplication
- shared string indexing
- relationship graph resolution

### 18.2 Golden tests
Test emitted XLSX package and XML structure:
- minimal workbook
- workbook with one worksheet
- basic cells
- shared strings
- styles
- merge cells
- 1900 / 1904 date system

### 18.3 Round-trip tests
- write with this library
- read with this library
- compare internal model / public behavior

### 18.4 Malformed file tests
Test damaged or non-ideal inputs:
- missing relationships
- missing parts
- invalid dimensions
- unsorted rows/cells
- style index out of range
- shared string index out of range
- unsupported extension parts
- extra unknown nodes

### 18.5 Compatibility tests
Test public API and behavior alignment:
- Aspose-style API smoke tests
- exception mapping tests
- file/stream parity tests
- behavior compatibility tests

### 18.6 Regression tests
Must be run for every feature expansion:
- XML structure drift
- style index stability
- shared string index stability
- save-open-save drift

---

## 19. Suggested Milestones

### Milestone 0 â€?contract freeze
Produce:
- `product_scope.md`
- `public_api.md`
- `exception_compatibility.md`
- `recovery_design.md`
- `styles.yaml`
- `dates.yaml`

### Milestone 1 â€?minimal workbook loop
Implement:
- create workbook
- add worksheet
- write `A1`
- save to file
- save to stream
- load from file
- load from stream

### Milestone 2 â€?cell values, shared strings, dates
Implement:
- text / number / bool / date
- shared strings
- inline strings
- 1900 / 1904 handling

### Milestone 3 â€?style core
Implement:
- full style API surface
- core style persistence
- style deduplication
- style round-trip

### Milestone 4 â€?recovery loader
Implement:
- diagnostics
- recoverable rules
- malformed tests

### Milestone 5 â€?commercial baseline
Implement:
- compatibility tests
- exception mapping
- validation before save
- samples and docs

---

## 20. Recommended ADR List

At minimum, keep these architecture decisions documented:

- `ADR-001 Public API mirrors Aspose.Cells`
- `ADR-002 Production code has no Open XML SDK dependency`
- `ADR-003 Sparse worksheet storage`
- `ADR-004 Style pool with copy-on-write values`
- `ADR-005 Recovery-friendly loader`
- `ADR-006 netstandard2.0 + net8.0 multi-targeting`
- `ADR-007 1900 and 1904 date systems supported from v0.1`
- `ADR-008 Formula persistence without calculation engine in v0.1`

---

## 21. Recommended First Batch of Deliverables

The first engineering batch should contain:
1. `docs/product_scope.md`
2. `docs/architecture/architecture.md`
3. `docs/api/public_api.md`
4. `docs/api/exception_compatibility.md`
5. `docs/architecture/recovery_design.md`
6. `docs/specs/opc/package.yaml`
7. `docs/specs/spreadsheetml/workbook.yaml`
8. `docs/specs/spreadsheetml/worksheet.yaml`
9. `docs/specs/spreadsheetml/styles.yaml`
10. `docs/specs/spreadsheetml/dates.yaml`
11. solution skeleton
12. project file skeletons
13. first class inventory
14. first test project skeleton

---

## 22. Recommended Immediate Next Step

The best next move is to produce the following **together**:

- the core markdown documents
- the YAML spec files
- the solution/project skeleton

That gives the Spec Agent, Architecture Agent, and Code Generator Agent a common, stable contract to work from.

---

## 23. Summary

This project should be executed as a **spec-driven, compatibility-aware, recovery-friendly engineering system** rather than a simple code-generation task.

The core strategic direction is:

- **Aspose-like public object model**
- **fully self-developed OPC + SpreadsheetML implementation**
- **commercial-grade robustness**
- **file + stream support from v0.1**
- **1900 / 1904 support from v0.1**
- **full `Style` API surface from v0.1**
- **repair-capable loading with diagnostics**
- **structured specs + layered architecture + automated validation**

