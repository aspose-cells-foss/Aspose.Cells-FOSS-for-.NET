# Aspose.Cells FOSS for .NET

A powerful, open-source .NET library for creating, editing, and saving Excel `.xlsx` workbooks. No Microsoft Excel installation required. Built with performance and simplicity in mind.

## ⭐ Why Aspose.Cells FOSS?

- 🚀 **Zero Dependencies** - Works without Microsoft Excel or Office
- 🎯 **Simple API** - Intuitive design modeled after Aspose.Cells
- ⚡ **High Performance** - Optimized for large workbooks
- 🛡️ **Reliable** - Recovery-oriented loading with detailed diagnostics
- 🔧 **Full Feature Set** - Cells, styles, charts, formulas, and more
- 📦 **Cross-Platform** - Supports .NET Standard 2.0 and .NET 8.0
- 🆓 **MIT License** - Free for commercial and personal use

## 🎯 Quick Start

### Installation

```bash
dotnet add package Aspose.Cells.FOSS
```

### Create Your First Excel File

```csharp
using Aspose.Cells_FOSS;

// Create a new workbook
var workbook = new Workbook();
var sheet = workbook.Worksheets[0];

// Add data
sheet.Cells["A1"].PutValue("Product");
sheet.Cells["B1"].PutValue("Price");
sheet.Cells["A2"].PutValue("Apple");
sheet.Cells["B2"].PutValue(2.99);
sheet.Cells["A3"].PutValue("Orange");
sheet.Cells["B3"].PutValue(1.99);

// Add formula
sheet.Cells["B4"].Formula = "=SUM(B2:B3)";

// Style the header
var headerStyle = sheet.Cells["A1"].GetStyle();
headerStyle.Font.IsBold = true;
headerStyle.Font.Color = System.Drawing.Color.White;
headerStyle.ForegroundColor = System.Drawing.Color.FromArgb(0, 120, 212);
sheet.Cells["A1"].SetStyle(headerStyle);
sheet.Cells["B1"].SetStyle(headerStyle);

// Save the workbook
workbook.Save("products.xlsx");
```

### Load and Edit Existing Files

```csharp
using Aspose.Cells_FOSS;

// Load an existing workbook
var workbook = new Workbook("existing.xlsx");
var sheet = workbook.Worksheets[0];

// Edit cells
sheet.Cells["A1"].PutValue("Updated Value");

// Add conditional formatting
var collection = sheet.ConditionalFormattings;
var index = collection.Add();
var format = collection[index];

var fcs = format.FormulaConditions;
fcs.AddCondition(FormatConditionType.Expression, OperatorType.Between, "=B2>100", "");

// Save changes
workbook.Save("updated.xlsx");
```

## ✨ Key Features

### Core Excel Operations
- ✅ Create, load, edit, and save `.xlsx` workbooks
- ✅ Support for file and stream operations
- ✅ Cell access by A1 notation and zero-based indexes
- ✅ Scalar values: string, number, boolean, decimal, DateTime
- ✅ Formula text storage and evaluation
- ✅ Shared strings and inline strings

### Styling & Formatting
- ✅ Rich cell styles (fonts, colors, backgrounds)
- ✅ Borders, fills, and number formats
- ✅ Cell merging
- ✅ Conditional formatting rules
- ✅ Data validation

### Worksheet Management
- ✅ Multiple worksheets with collection management
- ✅ Worksheet visibility control
- ✅ Auto-filter support
- ✅ Protection and security
- ✅ View settings and frozen panes

### Advanced Features
- ✅ Hyperlinks and defined names
- ✅ Data validations and rules
- ✅ Page setup, margins, headers/footers
- ✅ Print settings
- ✅ 1900 and 1904 date systems
- ✅ Workbook and document properties
- ✅ Recovery-oriented loading with diagnostics

## 📦 Compatibility

| Target Framework | Support |
|-----------------|---------|
| .NET Standard 2.0 | ✅ |
| .NET 8.0 | ✅ |
| .NET 6.0 | ✅ |
| .NET Framework 4.6.1+ | ✅ |
| .NET Core 2.0+ | ✅ |

**Minimum Requirements:**
- C# 6.0 or higher
- No external dependencies

## 🏗️ Build from Source

```bash
# Clone the repository
git clone https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET.git
cd Aspose.Cells-FOSS-for-.NET

# Build the library
dotnet build src\Aspose.Cells_FOSS\Aspose.Cells_FOSS.csproj -c Release

# Run samples
dotnet run --project samples\Aspose.Cells_FOSS.Samples.Basic\Aspose.Cells_FOSS.Samples.Basic.csproj
```

## 📚 Samples

Explore comprehensive examples in the [`samples/`](samples/) directory:

| Sample | Description |
|--------|-------------|
| [Basic](samples/Aspose.Cells_FOSS.Samples.Basic/) | Core operations and cell manipulation |
| [Loading](samples/Aspose.Cells_FOSS.Samples.Loading/) | Load options and diagnostics |
| [Styles](samples/Aspose.Cells_FOSS.Samples.Styles/) | Cell styling and formatting |
| [WorksheetSettings](samples/Aspose.Cells_FOSS.Samples.WorksheetSettings/) | Worksheet configuration |
| [Validations](samples/Aspose.Cells_FOSS.Samples.Validations/) | Data validation rules |
| [ConditionalFormatting](samples/Aspose.Cells_FOSS.Samples.ConditionalFormatting/) | Conditional formatting rules |
| [HyperlinksAndNames](samples/Aspose.Cells_FOSS.Samples.HyperlinksAndNames/) | Hyperlinks and defined names |
| [PageSetup](samples/Aspose.Cells_FOSS.Samples.PageSetup/) | Print and page setup |
| [Shapes](samples/Aspose.Cells_FOSS.Samples.Shapes/) | Drawing shapes |
| [Charts](samples/Aspose.Cells_FOSS.Samples.Charts/) | Chart creation |
| [Comments](samples/Aspose.Cells_FOSS.Samples.Comments/) | Cell comments |
| [DocumentProperties](samples/Aspose.Cells_FOSS.Samples.DocumentProperties/) | Workbook properties |
| [ListObjects](samples/Aspose.Cells_FOSS.Samples.ListObjects/) | Tables and lists |
| [Pictures](samples/Aspose.Cells_FOSS.Samples.Pictures/) | Image insertion |

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](License/LICENSE.txt) file for details.



## 📞 Support & Community

- 🐛 [Report Issues](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET/issues)
- 💬 [Discussions](https://github.com/aspose-cells-foss/Aspose.Cells-FOSS-for-.NET/discussions)
- 📧 [Email Support](mailto:support@aspose.com)



---

**Made with ❤️ by the Aspose.Cells FOSS Team**

If you find this project useful, please consider giving it a ⭐ on GitHub!
