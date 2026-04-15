# Aspose.Cells_FOSS

`Aspose.Cells_FOSS` is a .NET library for creating, loading, editing, and saving Excel `.xlsx` workbooks with an API shaped to stay close to Aspose.Cells for the features implemented in this repository.

This checkout contains the core library, runnable console samples, the solution file, and the project license.

## Targets And Requirements

- the library targets `netstandard2.0` and `net8.0`
- the project is compiled with `LangVersion` set to `6`
- consuming applications should be compatible with `netstandard2.0` or `net8.0`
- no Microsoft Excel installation is required

## Implemented Areas

Implemented feature areas in this checkout include:

- workbook create, load, and save from file and stream
- worksheet collection management
- cell access by A1 name and zero-based row and column indexes
- scalar cell values including string, number, boolean, decimal, `DateTime`, and formula text
- shared strings and inline strings
- styles, fonts, borders, fills, and number formats
- merges, hyperlinks, defined names, data validations, and conditional formatting
- row and column metadata
- worksheet visibility, worksheet view settings, protection, and auto filter support
- workbook and document properties
- 1900 and 1904 date systems
- recovery-oriented loading with diagnostics
- page setup, margins, headers and footers, and print settings

## Build

The verified build path for this checkout is to build project files directly from the repository root.

Build the library:

```powershell
dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Debug
```

Build a sample:

```powershell
dotnet build samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj -c Debug
```

The root solution file [`Aspose.Cells_FOSS.sln`](Aspose.Cells_FOSS.sln) is included for IDE use and groups the library plus the sample projects.

## Quick Start

```csharp
using Aspose.Cells_FOSS;

var workbook = new Workbook();
var sheet = workbook.Worksheets[0];

sheet.Cells["A1"].PutValue("Hello");
sheet.Cells["B1"].PutValue(123);
sheet.Cells["C1"].PutValue(true);
sheet.Cells["D1"].PutValue(new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc));
sheet.Cells["E1"].Formula = "=B1*2";

workbook.Save("hello.xlsx");

var loaded = new Workbook("hello.xlsx");
Console.WriteLine(loaded.Worksheets[0].Cells["E1"].Formula);
Console.WriteLine(loaded.Worksheets[0].Cells["A1"].StringValue);
```

## Samples

Runnable console samples live under [`samples/`](samples/README.md):

- `Aspose.Cells_FOSS.Samples.Basic`
- `Aspose.Cells_FOSS.Samples.Loading`
- `Aspose.Cells_FOSS.Samples.Styles`
- `Aspose.Cells_FOSS.Samples.WorksheetSettings`
- `Aspose.Cells_FOSS.Samples.Validations`
- `Aspose.Cells_FOSS.Samples.ConditionalFormatting`
- `Aspose.Cells_FOSS.Samples.HyperlinksAndNames`
- `Aspose.Cells_FOSS.Samples.PageSetup`

Run a sample from the repository root:

```powershell
dotnet run --project samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj -c Debug
```

## Repository Layout

- [`src/`](src): library source code
- [`samples/`](samples): runnable feature samples
- [`License/`](License): license text

This checkout does not include a `Spec/` directory or a `tests/` directory.

## License

This repository includes the MIT license at [`License/LICENSE.txt`](License/LICENSE.txt).

## Support

For bug reports, feature requests, and project questions, use the GitHub issue tracker:

- [Project issues](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET/issues)

