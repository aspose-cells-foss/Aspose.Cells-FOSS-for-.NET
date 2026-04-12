# Aspose.Cells_FOSS

`Aspose.Cells_FOSS` is a .NET library for creating, loading, editing, and saving Excel `.xlsx` workbooks with an API shaped to stay close to Aspose.Cells for the features implemented here.

The library currently targets `netstandard2.0` and `net8.0`. This checkout contains the core library source, runnable sample projects, and the project license.

## Feature Highlights

Implemented areas in this checkout include:

- workbook create, load, and save from file and stream
- worksheet collection management
- cell access by A1 name and zero-based row and column indexes
- scalar cell values including string, number, boolean, decimal, `DateTime`, and formula text
- shared strings and inline strings
- styles, fonts, borders, fills, and number formats
- merges, hyperlinks, defined names, data validations, and conditional formatting
- row and column metadata
- worksheet visibility, view settings, protection, and auto filter support
- workbook and document properties
- 1900 and 1904 date systems
- recovery-oriented loading with diagnostics
- page setup, margins, headers and footers, and print settings


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

Run a sample:

```powershell
dotnet run --project samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj -c Debug
```

## Solution And Build Notes

The root solution file is [`Aspose.Cells_FOSS.sln`](Aspose.Cells_FOSS.sln). In this checkout, the solution still references test projects under `tests\`, but that folder is not present.

If solution-level restore or build fails because of those missing projects, build the library and sample project files directly:

```powershell
dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Debug
dotnet build samples\Aspose.Cells_FOSS.Samples.PageSetup\Aspose.Cells_FOSS.Samples.PageSetup.csproj -c Debug
```

## Repository Layout

- [`src/`](src): library source code
- [`samples/`](samples): runnable feature samples
- [`License/`](License): license text

## License

This repository includes the MIT license at [`License/LICENSE.txt`](License/LICENSE.txt).
