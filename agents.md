# AGENTS.md

## Purpose

This file defines how agents should use the documents under `Spec/` to produce the final code for `Aspose.Cells_FOSS`.

The project is spec-driven. Code generation must follow the specs first, then the current source layout, and finally the tests.

## Source Of Truth

Read the following files before changing production code:

1. `Spec/product_scope.md`
2. `Spec/public_api.md`
3. `Spec/implementation_rules.md`
4. `Spec/Aspose.Cells_FOSS_for_DotNET_Agent_Architecture_v0.2.md`
5. `Spec/recovery_design.md`
6. `Spec/recovery_policy.yaml`
7. `Spec/step.md`
8. `Spec/api_compatibility_alignment.md`

Use the YAML specs as feature-level contracts:

- `Spec/opc_package.yaml`
- `Spec/workbook.yaml`
- `Spec/worksheet.yaml`
- `Spec/cells.yaml`
- `Spec/rows.yaml`
- `Spec/columns.yaml`
- `Spec/merges.yaml`
- `Spec/shared_strings.yaml`
- `Spec/hyperlinks.yaml`
- `Spec/data_validations.yaml`
- `Spec/conditional_formatting.yaml`
- `Spec/styles.yaml`
- `Spec/formulas.yaml`
- `Spec/dates.yaml`

## Priority Rules

When multiple files describe the same behavior, apply this priority:

1. `Spec/implementation_rules.md`
2. `Spec/public_api.md`
3. Feature YAML specs under `Spec/*.yaml`
4. `Spec/recovery_policy.yaml` and `Spec/recovery_design.md`
5. `Spec/Aspose.Cells_FOSS_for_DotNET_Agent_Architecture_v0.2.md`
6. `Spec/step.md`

If a conflict still exists, prefer:

- public behavior compatible with Aspose.Cells style
- recovery-friendly loading with explicit diagnostics
- deterministic XLSX serialization
- simple, Java-portable implementation style

## Mandatory Coding Rules

All generated code must follow `Spec/implementation_rules.md`:

- do not use `partial` classes
- keep each production `.cs` file under 800 lines
- do not use the `=>` operator
- do not use alias `using` directives
- prefer explicit loops and simple control flow

Also follow these project-level rules from the specs:

- production code must not depend on Open XML SDK
- support both file and stream load/save
- support both 1900 and 1904 date systems
- treat recovery and diagnostics as first-class behavior
- preserve deterministic output order for rows, cells, relationships, styles, and shared strings

## Implementation Workflow

For every feature or code change:

1. Read the relevant global docs in the Source Of Truth section.
2. Read the matching feature YAML files.
3. Map public API behavior from `Spec/public_api.md` to the internal model.
4. Implement load behavior, save behavior, validation, and recovery together.
5. Add or update tests for unit, golden, malformed, and compatibility coverage when applicable.
6. Verify the implementation against the spec rules, not only against current code shape.

Do not implement features that are outside `Spec/product_scope.md` v0.1 unless the user explicitly asks for them.

## Feature To Spec Mapping

### Workbook and package structure

Use:

- `Spec/product_scope.md`
- `Spec/public_api.md`
- `Spec/workbook.yaml`
- `Spec/opc_package.yaml`
- `Spec/recovery_policy.yaml`

Responsible for:

- `Workbook`
- `WorksheetCollection`
- workbook relationships
- workbook settings
- package part registration
- load/save entry points

### Worksheet grid and metadata

Use:

- `Spec/worksheet.yaml`
- `Spec/rows.yaml`
- `Spec/columns.yaml`
- `Spec/merges.yaml`
- `Spec/hyperlinks.yaml`
- `Spec/data_validations.yaml`
- `Spec/conditional_formatting.yaml`
- `Spec/recovery_policy.yaml`

Responsible for:

- worksheet XML
- dimension
- row metadata
- column metadata
- merge regions
- sheet ordering and visibility
- worksheet hyperlinks
- worksheet data validations
- worksheet conditional formatting

### Cell values and addressing

Use:

- `Spec/cells.yaml`
- `Spec/formulas.yaml`
- `Spec/dates.yaml`
- `Spec/shared_strings.yaml`
- `Spec/worksheet.yaml`

Responsible for:

- A1 parsing and formatting
- zero-based public indexers
- scalar values
- formula persistence
- blank cell rules
- date serial conversion
- shared string and inline string behavior

### Styles

Use:

- `Spec/styles.yaml`
- `Spec/dates.yaml`
- `Spec/public_api.md`
- `Spec/Aspose.Cells_FOSS_for_DotNET_Agent_Architecture_v0.2.md`

Responsible for:

- `Style`, `Font`, `Borders`, `Border`
- style repositories and deduplication
- style index allocation
- number format behavior for dates

### Recovery and diagnostics

Use:

- `Spec/recovery_design.md`
- `Spec/recovery_policy.yaml`
- all feature YAML `recovery` sections

Responsible for:

- fatal vs recoverable vs lossy recoverable behavior
- `LoadDiagnostics`
- warning callback integration
- repair rules
- data loss risk reporting

## Execution Order

Follow `Spec/step.md` as the preferred delivery order:

1. cell data APIs and Excel cell data import/export
2. style APIs and full style import/export
3. worksheet options and settings APIs and import/export
4. page setup APIs and import/export

Within each step, finish the whole vertical slice:

- public API
- internal model
- XML/package load
- XML/package save
- recovery rules
- tests

## Definition Of Done

A feature is complete only when all of the following are true:

- public API matches the intended shape in `Spec/public_api.md`
- implementation respects `Spec/implementation_rules.md`
- XML load/save behavior matches the relevant YAML specs
- recovery behavior matches the relevant `recovery` section
- deterministic serialization is preserved
- unsupported v0.1 features are not silently half-implemented
- tests cover nominal, round-trip, and malformed cases where required by the spec

## Anti-Patterns

Do not:

- invent behavior that contradicts the spec files
- spread date conversion logic outside a single date conversion service
- mix public facade objects with internal storage records
- rely on current incidental behavior when the spec says otherwise
- silently drop recoverable or lossy cases without diagnostics
- add unsupported advanced Excel features as hidden partial implementations

## Expected Output Style

Agents should produce code that is:

- small, explicit, and maintainable
- deterministic on save
- strict on fatal corruption
- tolerant on recoverable corruption
- easy to port to Java later
- aligned with the existing folder structure under `src/`, `tests/`, and `samples/`


