# Aspose.Cells FOSS for .NET

[![NuGet version](https://img.shields.io/nuget/v/Aspose.Cells.FOSS.svg)](https://www.nuget.org/packages/Aspose.Cells.FOSS/) [![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET/blob/master/License/LICENSE.txt) [![Contributors](https://img.shields.io/github/contributors/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET.svg)](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET/graphs/contributors) [![Issues](https://img.shields.io/github/issues/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET.svg)](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET/issues)

[![Aspose.Cells FOSS for .NET](https://products.aspose.org/media/cells/net/banner-readme.png)](https://products.aspose.org/cells/net/)

Aspose.Cells FOSS for .NET is an MIT-licensed spreadsheet library for .NET. It provides an Aspose.Cells-style API for creating, loading, editing, and saving Excel `.xlsx` workbooks, and it can also export supported workbook content to PDF through the built-in renderer.

The package targets `net8.0` and requires .NET 8 or later. It runs on Windows, Linux, and macOS, and does not require Microsoft Excel or any native Office installation. The NuGet package id is `Aspose.Cells.FOSS`; the public namespace is `Aspose.Cells_FOSS`.

The `netstandard2.0` target has been removed. Applications targeting .NET Framework, .NET Core, or .NET 5-7 must upgrade to .NET 8 or later, or remain on a previous compatible package version.

## Navigation

- [Key Capabilities](#key-capabilities)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Examples](#examples)
- [Documentation](#documentation)
- [Scope and Limitations](#scope-and-limitations)
- [Development and Testing](#development-and-testing)
- [License](#license)

## Key Capabilities

- Create, load, modify, and save Excel `.xlsx` workbooks with `Workbook`, `Worksheet`, `Cells`, and `Cell`.
- Load malformed packages with recovery-oriented `LoadOptions` and inspect `LoadDiagnostics`.
- Read and write scalar values, formulas, dates, and number formats.
- Apply styles including fonts, fills, borders, alignment, wrapping, and locking.
- Configure validations, conditional formatting, hyperlinks, defined names, filters, tables, comments, pictures, shapes, charts, and page setup.
- Export workbooks to PDF with `SaveFormat.Pdf` or `PdfSaveOptions`.

## Installation

```bash
dotnet add package Aspose.Cells.FOSS
```

```xml
<PackageReference Include="Aspose.Cells.FOSS" Version="26.10.0.0" />
```

## Quick Start

```csharp
using Aspose.Cells_FOSS;

var workbook = new Workbook();
var sheet = workbook.Worksheets[0];

sheet.Name = "Products";
sheet.Cells["A1"].PutValue("Product");
sheet.Cells["B1"].PutValue("Price");
sheet.Cells["A2"].PutValue("Apple");
sheet.Cells["B2"].PutValue(2.99m);
sheet.Cells["A3"].PutValue("Orange");
sheet.Cells["B3"].PutValue(1.99m);
sheet.Cells["B4"].Formula = "=SUM(B2:B3)";

var headerStyle = sheet.Cells["A1"].GetStyle();
headerStyle.Font.IsBold = true;
headerStyle.Font.Color = Color.FromArgb(255, 255, 255, 255);
headerStyle.Pattern = FillPattern.Solid;
headerStyle.ForegroundColor = Color.FromArgb(255, 34, 120, 212);

sheet.Cells["A1"].SetStyle(headerStyle);
sheet.Cells["B1"].SetStyle(headerStyle);

workbook.Save("products.xlsx");
workbook.Save("products.pdf", new PdfSaveOptions());
```

Load an existing workbook with repair-friendly options:

```csharp
using System;
using Aspose.Cells_FOSS;

var loadOptions = new LoadOptions
{
    TryRepairPackage = true,
    TryRepairXml = true,
    StrictMode = false
};

var workbook = new Workbook("input.xlsx", loadOptions);

if (workbook.LoadDiagnostics.HasDataLossRisk)
{
    Console.WriteLine("Potential data loss risk detected during load.");
}

workbook.Worksheets[0].Cells["A1"].PutValue("Updated");
workbook.Save("updated.xlsx");
```

## Examples

Runnable samples live under [`samples/`](samples/README.md). Feature areas currently covered include:

- basic workbook and cell operations
- load and repair behavior
- styles
- worksheet settings
- validations
- conditional formatting
- hyperlinks and defined names
- page setup
- comments
- charts
- list objects
- pictures
- document properties
- shapes
- PDF conversion

Run one sample from the repository root:

```bash
dotnet run --project samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj
```

## Documentation

- [Local documentation landing page](documentation/_site/index.html)
- [Samples overview](samples/README.md)
- [V26.10.0 release notes](RELEASE_NOTES_V26.10.0.md)
- [Contributor and agent instructions](agents.md)

## Scope and Limitations

- Only `.xlsx` is supported as the spreadsheet file format for load and save.
- The library stores formulas but does not include a formula calculation engine.
- PDF export is supported, but there is no general HTML export or image export API.
- Macros, VBA, and pivot tables are outside the supported public API surface.
- Broader format coverage and richer spreadsheet features belong to the commercial Aspose.Cells product line.

## Development and Testing

Building from source requires the .NET 8 SDK or a newer SDK capable of targeting `net8.0`. The library and samples target `net8.0`.

Build the main library:

```bash
dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Debug
```

Build the basic sample:

```bash
dotnet build samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj -c Debug
```

This checkout does not contain a `tests/` directory. These builds check compilation; they do not provide automated behavioral test coverage.

## License

This project is licensed under the [MIT License](License/LICENSE.txt).
