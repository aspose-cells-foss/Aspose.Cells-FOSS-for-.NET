# AGENTS.md

## Purpose

This file defines how agents should work in this checkout of `Aspose.Cells_FOSS`.

The repository is currently source-driven. Agents must align changes to the code and assets that are present in this checkout instead of assuming missing spec or test folders exist.

## Current Checkout Facts

Before making changes, account for the current repository state:

- the repository contains `src/`, `samples/`, `License/`, `documentation/_site/`, the root solution file, and this document (`agents.md`)
- there is no `Spec/` directory in this checkout
- there is no `tests/` directory in this checkout
- the library and sample projects target `net8.0`; .NET 8 or later is required, and `netstandard2.0` is no longer targeted
- the library is compiled with `LangVersion` set to `6`
- the current project version is `26.10.0.0` (normalized NuGet version `26.10.0`); the package ID is `Aspose.Cells.FOSS` and the public namespace is `Aspose.Cells_FOSS`
- current release notes are in `RELEASE_NOTES_V26.10.0.md`, currently marked unreleased; prior notes include V26.9.0
- the PDF renderer uses SkiaSharp; XLSX input and XLSX/PDF output are supported, with no general formula calculation engine
- `Aspose.Cells_FOSS.sln` includes the library and all 15 sample projects, including `Aspose.Cells_FOSS.Samples.PdfConversion` under the `samples` solution folder
- `documentation/` currently contains generated `_site/` output, not a documentation source project or build guide

Recheck these facts against the working tree before changing them. Do not write instructions or code that depend on files that are not present unless the user explicitly asks you to add them.

## Source Of Truth

When changing code or docs in this checkout, use the following priority:

1. direct user instructions
2. this `agents.md`
3. the current public API and behavior in `src/Aspose.Cells_FOSS/`
4. runnable examples under `samples/`
5. the root `README.md`, solution file, and project files

If a future checkout adds a `Spec/` directory, treat those specs as the feature contract for the areas they cover. In this checkout, do not block work on missing specs.

## Working Rules

All code changes should preserve the existing implementation constraints that are visible in the repo:

- keep the library compatible with C# 6 syntax
- preserve the current `net8.0` target unless the user explicitly asks to change targets; the newer runtime does not remove the C# 6 language constraint
- prefer explicit control flow and simple, maintainable code
- keep save output deterministic where current serializer code depends on ordering
- preserve file and stream load and save support
- treat load diagnostics and recovery behavior as important public behavior
- do not introduce a dependency on Microsoft Excel

Current API alignment expectations for this checkout (V26.10.0):

- preserve existing Aspose-compatible overloads and names
- keep `FontUnderlineType`, `StyleFlag`, and selective style-application behavior intact
- keep `CellValueType` and `CellArea` in their Aspose-aligned shapes
- keep `Style.Font` and `Style.Borders` non-assignable at object level (`private set`)
- preserve rich-text APIs (`Cell.Characters`, `GetCharacters`, `SetCharacters`, `IsRichText`, and `FontSetting`) and their load/save behavior
- preserve `SaveFormat.Pdf`, `PdfSaveOptions`, and automatic PDF selection for `Workbook.Save(string)` paths ending in `.pdf`
- preserve `PdfSaveOptions.ConditionalFormattingEvaluationDate` semantics: a supplied date overrides the local export date for relative-date conditional formatting
- treat release-noted behavior changes as public contract unless the user explicitly requests a change

When editing existing files:

- read the surrounding code first and follow the local style
- do not revert unrelated work already present in the worktree
- avoid broad refactors unless they are required for the requested change
- update docs and samples when the public behavior they describe changes

## Implementation Workflow

For feature work or bug fixes:

1. inspect the relevant public API types under `src/Aspose.Cells_FOSS/`
2. inspect the matching internal model, root-level `XlsxWorkbook*` helpers, packaging, and XML mapper code; for PDF changes, also inspect `Rendering/`
3. inspect any related sample projects under `samples/`
4. implement the affected public behavior, load path, save path, rendering path, and diagnostics together; distinguish data preserved for rendering from data serialized back to XLSX
5. run the narrowest practical verification command for the touched project or sample
6. update `README.md`, `samples/README.md`, release notes, or sample code if user-facing behavior changed

Because `tests/` is absent in this checkout, do not claim automated test coverage that you did not run. Prefer targeted `dotnet build` verification and clearly state what was and was not validated.

## Repository Map

Use the current layout as the implementation map:

- `src/Aspose.Cells_FOSS/`: public API surface and XLSX implementation
- `src/Aspose.Cells_FOSS/Core/`: internal workbook, worksheet, style, and value models
- `src/Aspose.Cells_FOSS/Packaging/`: OPC package abstractions and relationship handling
- `src/Aspose.Cells_FOSS/Xml/`: XML load and save mappers
- `src/Aspose.Cells_FOSS/Rendering/`: PDF export, worksheet layout and pagination, fonts, charts, shapes, images, and conditional-format evaluation
- `src/Aspose.Cells_FOSS/Validation/`: workbook validation messages and validator
- `samples/`: runnable console samples for implemented features
  - `Aspose.Cells_FOSS.Samples.Basic`: core operations and cell manipulation
  - `Aspose.Cells_FOSS.Samples.Loading`: load options and diagnostics
  - `Aspose.Cells_FOSS.Samples.Styles`: cell styling and formatting
  - `Aspose.Cells_FOSS.Samples.WorksheetSettings`: worksheet configuration
  - `Aspose.Cells_FOSS.Samples.Validations`: data validation rules
  - `Aspose.Cells_FOSS.Samples.ConditionalFormatting`: conditional formatting rules
  - `Aspose.Cells_FOSS.Samples.HyperlinksAndNames`: hyperlinks and defined names
  - `Aspose.Cells_FOSS.Samples.PageSetup`: print and page setup
  - `Aspose.Cells_FOSS.Samples.Shapes`: drawing shapes
  - `Aspose.Cells_FOSS.Samples.Charts`: chart creation
  - `Aspose.Cells_FOSS.Samples.Comments`: cell comments
  - `Aspose.Cells_FOSS.Samples.DocumentProperties`: workbook properties
  - `Aspose.Cells_FOSS.Samples.ListObjects`: tables and lists
  - `Aspose.Cells_FOSS.Samples.Pictures`: image insertion
  - `Aspose.Cells_FOSS.Samples.PdfConversion`: XLSX-to-PDF conversion with optional evaluation-date and print overrides
- `samples/files/`: workbook inputs used by loading and PDF conversion samples
- `documentation/_site/`: generated documentation; do not assume its contents describe the latest source
- `License/`: license files

## Build And Verification

Use the .NET 8 SDK or a newer SDK capable of targeting `net8.0`. Run commands from the repository root and prefer targeted project builds over broad solution commands.

Recommended commands:

- `dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Debug`
- `dotnet build samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj -c Debug`
- `dotnet build samples\Aspose.Cells_FOSS.Samples.PdfConversion\Aspose.Cells_FOSS.Samples.PdfConversion.csproj -c Debug`

Use additional sample project builds when your change affects that area. Use `--no-restore` only when the required dependency assets already exist. Report dependency-restore failures separately from compilation failures.

For PDF rendering changes:

- select a relevant existing XLSX file and pass its path after `--` to `dotnet run --project samples\Aspose.Cells_FOSS.Samples.PdfConversion\Aspose.Cells_FOSS.Samples.PdfConversion.csproj`
- the sample locates the repository through `samples/files/`, even when an explicit input path is supplied; without input arguments it recursively converts all XLSX files there
- output is written beside each input with the same basename and a `.pdf` extension; use a working copy when a matching PDF is a reference artifact
- use `--conditional-date yyyy-MM-dd` for reproducible relative-date conditional formatting; optional print overrides are `--paper-size A4`, `--scale 100`, and the four `--margin-*-cm` options
- compare affected pages with available reference output, checking pagination, clipping, text, and drawing placement; record relevant font, culture, and print settings
- distinguish successful compilation, successful conversion, and visual comparison in the result; none implies the others

For documentation-only changes, check referenced paths, project settings, commands, and the diff; a new build is not required unless the change raises a code or build concern.

## Documentation Rules

When updating docs:

- describe only behavior that exists in this checkout
- avoid references to missing `Spec/` or `tests/` folders unless you explicitly call out that they are absent
- keep build instructions executable from the repository root
- prefer short, concrete examples over broad feature claims that cannot be verified locally
- when behavior changes, update `RELEASE_NOTES_V26.10.0.md` with affected API and compatibility notes; keep historical release notes intact unless correcting a verified historical error
- for V26.10.0 release reviews, compare with the V26.9.0 commit (`666f894`) and include relevant uncommitted and untracked source files; do not describe existing V26.9.0 features as new
- keep the project version, README package example, and package release-notes URL consistent when bumping versions; do not invent a release date or claim publication
- `.gitignore` currently ignores `RELEASE*.md`; check tracking status when preparing release-note changes for a commit
- document renderer limitations explicitly: supported chart rendering is not equivalent to new public chart APIs, and limited conditional-format expression evaluation is not a general formula engine
- do not infer complete XLSX round-trip preservation from successful loading or PDF rendering

## Anti-Patterns

Do not:

- assume spec files exist when they do not
- describe solution contents that are not in `Aspose.Cells_FOSS.sln`
- claim automated tests were updated when the repo has no test project in this checkout
- add modern C# syntax that conflicts with `LangVersion` 6
- silently change public API behavior without updating docs or samples
- claim PDF visual fidelity from a successful build alone

## Expected Output Style

Agents should produce work that is:

- accurate to the current checkout
- small and explicit
- compatible with the existing API shape
- straightforward to verify with targeted builds
- clear about validation limits when tests are unavailable
