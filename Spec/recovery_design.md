# Aspose.Cells FOSS for .NET — Recovery Design

This document describes how the library handles damaged XLSX files.

Goal: load **real-world XLSX files that are slightly corrupted but recoverable**.

## Error Categories

### Fatal Errors
Loading cannot continue.

Examples:
- File is not a ZIP archive
- Workbook part missing and cannot be inferred
- Package structure unreadable

Action:
Throw `WorkbookLoadException`.

---

### Recoverable Errors
File can be repaired automatically.

Examples:

- worksheet rows unsorted
- cell references unordered
- missing dimension element
- missing optional relationships

Action:

1. Record issue in diagnostics
2. Apply repair
3. Continue loading

---

### Lossy Recoverable Errors

Data can be loaded but may lose information.

Examples:

- style index out of range
- shared string index invalid
- unsupported extension parts
- unsupported future feature parts

Action:

1. Record warning
2. Preserve raw data if possible
3. Mark `HasDataLossRisk = true`

## Recovery Layers

### Package Layer

Repair package-level issues:

- infer missing content types
- infer missing relationships
- tolerate extra unknown parts

### XML Layer

Repair XML-level issues:

- reorder rows
- reorder cells
- ignore unknown elements
- normalize node ordering

### Semantic Layer

Repair logical issues:

- fallback default style
- fallback empty shared string
- drop invalid cached formula values

## Diagnostics

All issues should be recorded in `LoadDiagnostics`.

Example:

```csharp
var wb = new Workbook(stream, new LoadOptions
{
    TryRepairPackage = true
});

var diagnostics = wb.LoadDiagnostics;
```

Diagnostics include:

- issue type
- severity
- repair applied
- possible data loss