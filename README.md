# Aspose.Cells_FOSS

`Aspose.Cells_FOSS` is a spec-driven .NET library for creating, loading, editing, and saving Excel `.xlsx` workbooks with an API shaped to be close to Aspose.Cells for supported features.

The current implementation targets `netstandard2.0` and `net8.0` and focuses on deterministic XLSX serialization, recovery-friendly loading, and a small explicit codebase that can be ported more easily.

## Status

This repository is currently aligned to the v0.1 scope defined under [`Spec/`](Spec/product_scope.md).

Implemented areas include:

- workbook create, load, and save from file and stream
- worksheet collection management
- cell access by A1 name and row or column index
- scalar cell values: string, number, boolean, datetime, and formula
- shared strings and inline strings
- core style persistence
- merges, hyperlinks, data validations, and conditional formatting
- row and column metadata
- worksheet settings, protection, and auto filter baseline
- workbook properties, document properties, view settings, and defined names
- 1900 and 1904 date systems
- repair-oriented loading with diagnostics
- page setup and print settings

Not supported in v0.1:

- charts
- images and drawings
- pivot tables
- tables
- comments
- macros or VBA
- workbook encryption
- formula calculation engine
- streaming writer
- x14 conditional-formatting extensions beyond the main SpreadsheetML model

## Project Goals

- keep the public API close to Aspose.Cells where the feature is supported
- implement XLSX handling without depending on Open XML SDK in production code
- preserve deterministic output order during save
- treat load recovery and diagnostics as first-class behavior
- keep the implementation simple and portable

## Quick Start

Build the library:

```powershell
dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Debug
```

Basic usage:

```csharp
using Aspose.Cells_FOSS;

var workbook = new Workbook();
var sheet = workbook.Worksheets[0];

sheet.Cells["A1"].PutValue("Hello");
sheet.Cells["B1"].PutValue(123);
sheet.Cells["C1"].Formula = "=B1*2";

workbook.Save("hello.xlsx");

var loaded = new Workbook("hello.xlsx");
Console.WriteLine(loaded.Worksheets[0].Cells["C1"].Formula);
Console.WriteLine(loaded.Worksheets[0].Cells["C1"].StringValue);
```

## Samples

Runnable sample projects live under [`samples/`](samples/README.md):

- `Aspose.Cells_FOSS.Samples.Basic`
- `Aspose.Cells_FOSS.Samples.Loading`
- `Aspose.Cells_FOSS.Samples.Styles`
- `Aspose.Cells_FOSS.Samples.WorksheetSettings`
- `Aspose.Cells_FOSS.Samples.Validations`
- `Aspose.Cells_FOSS.Samples.ConditionalFormatting`
- `Aspose.Cells_FOSS.Samples.HyperlinksAndNames`
- `Aspose.Cells_FOSS.Samples.PageSetup`

Run a sample:

```powershell
dotnet run --project samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj -c Debug
```

## Build And Test

Build the library and sample projects with `dotnet build` on the individual project files.

Examples:

```powershell
dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Debug
dotnet build samples\Aspose.Cells_FOSS.Samples.PageSetup\Aspose.Cells_FOSS.Samples.PageSetup.csproj -c Debug
```

The repository also contains console-based test projects under [`tests/`](tests) for:

- unit coverage
- compatibility behavior
- malformed input handling
- golden round-trip verification
- OpenXML comparison and feature-focused interoperability checks

Example test runs:

```powershell
dotnet run --project tests\Aspose.Cells_FOSS.UnitTests\Aspose.Cells_FOSS.UnitTests.csproj -c Debug
dotnet run --project tests\Aspose.Cells_FOSS.CompatibilityTests\Aspose.Cells_FOSS.CompatibilityTests.csproj -c Debug
```

There are also helper scripts under [`tests/`](tests) for selected comparison and OpenXML scenarios.

## Repository Layout

- [`src/`](src): library source code
- [`samples/`](samples): runnable feature samples
- [`tests/`](tests): console-based test projects and shared test infrastructure
- [`Spec/`](Spec): product scope, public API, feature contracts, and implementation rules
- [`Input/`](Input): sample input workbooks used by tests and compatibility checks

## Specs And Development Model

This project is spec-driven. The primary design and behavior contracts live in [`Spec/`](Spec), especially:

- [`Spec/product_scope.md`](Spec/product_scope.md)
- [`Spec/public_api.md`](Spec/public_api.md)
- [`Spec/implementation_rules.md`](Spec/implementation_rules.md)
- feature YAML files such as [`Spec/workbook.yaml`](Spec/workbook.yaml), [`Spec/worksheet.yaml`](Spec/worksheet.yaml), and [`Spec/styles.yaml`](Spec/styles.yaml)

The implementation prioritizes spec compliance before existing code shape.

## Current Focus

The repository already includes vertical slices for:

- cell data APIs and XLSX import or export
- core style APIs and style XML persistence
- worksheet settings and metadata
- validations, hyperlinks, conditional formatting, and page setup
- workbook metadata, defined names, and recovery-oriented loading

## License

No root license file is present in this repository at the moment. Add one before publishing or redistributing the project.
