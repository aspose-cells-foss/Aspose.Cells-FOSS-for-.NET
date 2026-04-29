# AGENTS.md

## Purpose

This file defines how agents should work in this checkout of `Aspose.Cells_FOSS`.

The repository is currently source-driven. Agents must align changes to the code and assets that are present in this checkout instead of assuming missing spec or test folders exist.

## Current Checkout Facts

Before making changes, account for the current repository state:

- the repository contains `src/`, `samples/`, `License/`, the root solution file, and this document
- there is no `Spec/` directory in this checkout
- there is no `tests/` directory in this checkout
- `src/Aspose.Cells_FOSS/Aspose.Cells_FOSS.csproj` targets `netstandard2.0` and `net8.0`
- the library is compiled with `LangVersion` set to `6`

Do not write instructions or code that depend on files that are not present unless the user explicitly asks you to add them.

## Source Of Truth

When changing code or docs in this checkout, use the following priority:

1. direct user instructions
2. this `AGENTS.md`
3. the current public API and behavior in `src/Aspose.Cells_FOSS/`
4. runnable examples under `samples/`
5. the root `README.md`, solution file, and project files

If a future checkout adds a `Spec/` directory, treat those specs as the feature contract for the areas they cover. In this checkout, do not block work on missing specs.

## Working Rules

All code changes should preserve the existing implementation constraints that are visible in the repo:

- keep the library compatible with C# 6 syntax
- preserve `netstandard2.0` support unless the user explicitly asks to change targets
- prefer explicit control flow and simple, maintainable code
- keep save output deterministic where current serializer code depends on ordering
- preserve file and stream load and save support
- treat load diagnostics and recovery behavior as important public behavior
- do not introduce a dependency on Microsoft Excel

When editing existing files:

- read the surrounding code first and follow the local style
- do not revert unrelated work already present in the worktree
- avoid broad refactors unless they are required for the requested change
- update docs and samples when the public behavior they describe changes

## Implementation Workflow

For feature work or bug fixes:

1. inspect the relevant public API types under `src/Aspose.Cells_FOSS/`
2. inspect the matching internal model, packaging, and XML mapper code under `src/Aspose.Cells_FOSS/Core/`, `Packaging/`, and `Xml/`
3. inspect any related sample projects under `samples/`
4. implement the public behavior, load path, save path, and diagnostics together
5. run the narrowest practical verification command for the touched project or sample
6. update `README.md` or sample code if the user-facing workflow changed

Because `tests/` is absent in this checkout, do not claim automated test coverage that you did not run. Prefer targeted `dotnet build` verification and clearly state what was and was not validated.

## Repository Map

Use the current layout as the implementation map:

- `src/Aspose.Cells_FOSS/`: public API surface and XLSX implementation
- `src/Aspose.Cells_FOSS/Core/`: internal workbook, worksheet, style, and value models
- `src/Aspose.Cells_FOSS/Packaging/`: OPC package abstractions and relationship handling
- `src/Aspose.Cells_FOSS/Xml/`: XML load and save mappers
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
- `License/`: license files

## Build And Verification

Prefer targeted project builds over broad solution commands.

Recommended commands:

- `dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Debug`
- `dotnet build samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj -c Debug`

Use additional sample project builds when your change affects that area.

## Documentation Rules

When updating docs:

- describe only behavior that exists in this checkout
- avoid references to missing `Spec/` or `tests/` folders unless you explicitly call out that they are absent
- keep build instructions executable from the repository root
- prefer short, concrete examples over broad feature claims that cannot be verified locally

## Anti-Patterns

Do not:

- assume spec files exist when they do not
- describe solution contents that are not in `Aspose.Cells_FOSS.sln`
- claim automated tests were updated when the repo has no test project in this checkout
- add modern C# syntax that conflicts with `LangVersion` 6
- silently change public API behavior without updating docs or samples

## Expected Output Style

Agents should produce work that is:

- accurate to the current checkout
- small and explicit
- compatible with the existing API shape
- straightforward to verify with targeted builds
- clear about validation limits when tests are unavailable
