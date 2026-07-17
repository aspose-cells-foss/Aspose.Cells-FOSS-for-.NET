# Aspose.Cells FOSS for .NET

Aspose.Cells FOSS is an open-source .NET library for creating, loading, editing, and saving Excel `.xlsx` workbooks without requiring Microsoft Excel.

## Highlights

- Aspose.Cells-compatible API surface for common XLSX scenarios
- `.xlsx` load/save from file paths and streams
- Cell values, formulas, styles, merges, and number formats
- Conditional formatting, data validation, hyperlinks, names, tables, charts, shapes, comments, pictures
- Recovery-oriented loading with diagnostics and warning callbacks
- Multi-targeted library: `netstandard2.0` and `net8.0`

## Installation

```bash
dotnet add package Aspose.Cells.FOSS
```

## Product Links

- Open-source product page: <https://products.aspose.org/cells>
- Aspose.Cells product page: <https://products.aspose.com/cells>

## Quick Start

### Create, style, and save a workbook

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
```

### Load workbook with recovery options and diagnostics

```csharp
using System;
using Aspose.Cells_FOSS;

public sealed class ConsoleWarningCallback : IWarningCallback
{
    public void Warning(WarningInfo warningInfo)
    {
        Console.WriteLine("[{0}] {1}: {2}", warningInfo.Severity, warningInfo.Code, warningInfo.Message);
    }
}

var loadOptions = new LoadOptions
{
    TryRepairPackage = true,
    TryRepairXml = true,
    StrictMode = false,
    WarningCallback = new ConsoleWarningCallback()
};

var workbook = new Workbook("input.xlsx", loadOptions);

if (workbook.LoadDiagnostics.HasRepairs)
{
    Console.WriteLine("Load repairs were applied.");
}

if (workbook.LoadDiagnostics.HasDataLossRisk)
{
    Console.WriteLine("Potential data loss risk detected during load.");
}

workbook.Worksheets[0].Cells["A1"].PutValue("Updated");
workbook.Save("updated.xlsx");
```

## Supported Areas

- Workbook and worksheet management
- Cells and formulas
- Style model (font, fill, borders, alignment, number formats)
- Worksheet settings (visibility, zoom, RTL, gridlines, protection)
- Hyperlinks and defined names
- Data validation and conditional formatting
- Page setup and print-related settings
- Tables (`ListObjects`)
- Pictures, shapes, and charts
- Document properties

For runnable examples, see [`samples/`](samples/):

- `Aspose.Cells_FOSS.Samples.Basic`
- `Aspose.Cells_FOSS.Samples.Loading`
- `Aspose.Cells_FOSS.Samples.Styles`
- `Aspose.Cells_FOSS.Samples.WorksheetSettings`
- `Aspose.Cells_FOSS.Samples.Validations`
- `Aspose.Cells_FOSS.Samples.ConditionalFormatting`
- `Aspose.Cells_FOSS.Samples.HyperlinksAndNames`
- `Aspose.Cells_FOSS.Samples.PageSetup`
- `Aspose.Cells_FOSS.Samples.Shapes`
- `Aspose.Cells_FOSS.Samples.Charts`
- `Aspose.Cells_FOSS.Samples.Comments`
- `Aspose.Cells_FOSS.Samples.DocumentProperties`
- `Aspose.Cells_FOSS.Samples.ListObjects`
- `Aspose.Cells_FOSS.Samples.Pictures`

## Build from Source

Run commands from the repository root:

```bash
dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Debug
dotnet build samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj -c Debug
```

## Compatibility

Primary targets:

- `netstandard2.0`
- `net8.0`

Because the package targets `netstandard2.0`, it can also be consumed by many compatible runtimes (for example .NET Framework 4.6.1+ and modern .NET).

## License

MIT. See [License/LICENSE.txt](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET/blob/master/License/LICENSE.txt).

## Support

- Issues: <https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET/issues>
