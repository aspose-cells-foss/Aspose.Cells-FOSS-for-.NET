# Aspose.Cells FOSS for .NET - Product Scope v0.1

## Project Goal
Build a commercial-grade .NET library to manipulate Excel XLSX files based on ECMA-376 / SpreadsheetML / OPC while keeping a public API highly compatible with Aspose.Cells.

Namespace: `Aspose.Cells_FOSS`

Target frameworks:
- netstandard2.0
- net8.0

## Supported in v0.1

### Workbook
- Create workbook
- Load from file
- Load from stream
- Save to file
- Save to stream
- Workbook settings
- Worksheet collection
- Workbook defined names baseline

### Worksheets
- Add worksheet
- Remove worksheet
- Rename worksheet
- Access by index
- Access by name

### Cells
- Access using A1 notation
- Access using row/column indexes
- Read / write cell values
- Supported types:
  - string
  - number
  - boolean
  - datetime
  - formula

### Structural worksheet features
- merge cells
- hyperlinks
- data validation
- conditional formatting
- row height
- column width
- hidden rows
- hidden columns
- worksheet dimension
- sheet protection baseline
- auto filter baseline

### Shared Strings
- shared string table support
- inline strings support

### Styles (core persistence)
- font, including strikethrough
- fill, including classic pattern fills
- border, including diagonal border state and classic SpreadsheetML border styles
- number format, including built-in and custom format APIs
- horizontal alignment
- vertical alignment
- wrap text
- indent / text rotation / shrink-to-fit / reading order / relative indent
- locked / hidden
- differential formats for classic conditional formatting

### Date System
- 1900 system
- 1904 system

### Load robustness
- damaged but recoverable XLSX files supported
- strict loading mode
- repair-friendly loading mode
- load diagnostics

## Not Supported in v0.1

- charts
- images
- drawings
- pivot tables
- tables
- comments
- macros (VBA)
- encryption
- formula calculation engine
- streaming writer
- conditional formatting x14 extensions and advanced threshold/visual options beyond the main SpreadsheetML model

## Key Principles

1. Object model compatible with Aspose.Cells, prioritizing member names and signatures for supported v0.1 features
2. Fully self-developed XLSX implementation
3. Spec-driven development
4. Strong testing, validation, and OpenXML interoperability verification
5. Recovery-friendly loading





