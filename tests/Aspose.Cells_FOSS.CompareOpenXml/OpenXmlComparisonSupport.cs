using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static Aspose.Cells_FOSS.CompareOpenXml.ComparisonValueHelpers;

namespace Aspose.Cells_FOSS.CompareOpenXml;

internal static class OpenXmlComparisonSupport
{
    internal static Dictionary<string, CellSnapshot> ReadCellDataWithOpenXmlSdk(string workbookPath)
    {
        return ReadWithOpenXmlSdk(
            workbookPath,
            delegate(OpenXmlReadContext context)
            {
                var map = new Dictionary<string, CellSnapshot>(StringComparer.OrdinalIgnoreCase);
                foreach (var descriptor in EnumerateOpenXmlCells(context))
                {
                    map[BuildCellKey(descriptor.SheetName, descriptor.CellReference)] = CreateOpenXmlCellSnapshot(
                        descriptor.SheetName,
                        descriptor.CellReference,
                        descriptor.Cell,
                        context.SharedStrings,
                        context.Styles.DateStyleIndexes,
                        context.DateSystem);
                }

                return map;
            });
    }

    internal static Dictionary<string, StyleSnapshot> ReadCellStylesWithOpenXmlSdk(string workbookPath)
    {
        return ReadWithOpenXmlSdk(
            workbookPath,
            delegate(OpenXmlReadContext context)
            {
                var map = new Dictionary<string, StyleSnapshot>(StringComparer.OrdinalIgnoreCase);
                foreach (var descriptor in EnumerateOpenXmlCells(context))
                {
                    var snapshot = CreateOpenXmlStyleSnapshot(descriptor.Cell, context.Styles) with
                    {
                        SheetName = descriptor.SheetName,
                        CellReference = descriptor.CellReference,
                    };
                    map[BuildCellKey(descriptor.SheetName, descriptor.CellReference)] = snapshot;
                }

                return map;
            });
    }

    internal static StyleSnapshot CreateLibraryStyleSnapshot(string sheetName, string cellReference, Style style)
    {
        return new StyleSnapshot(
            sheetName,
            cellReference,
            style.Font.Name,
            NormalizeDouble(style.Font.Size),
            style.Font.Bold,
            style.Font.Italic,
            style.Font.Underline,
            style.Font.StrikeThrough,
            NormalizeColor(style.Font.Color),
            style.Pattern.ToString(),
            NormalizeColor(style.ForegroundColor),
            NormalizeColor(style.BackgroundColor),
            style.Borders.Left.LineStyle.ToString(),
            NormalizeColor(style.Borders.Left.Color),
            style.Borders.Right.LineStyle.ToString(),
            NormalizeColor(style.Borders.Right.Color),
            style.Borders.Top.LineStyle.ToString(),
            NormalizeColor(style.Borders.Top.Color),
            style.Borders.Bottom.LineStyle.ToString(),
            NormalizeColor(style.Borders.Bottom.Color),
            style.Borders.Diagonal.LineStyle.ToString(),
            NormalizeColor(style.Borders.Diagonal.Color),
            style.Borders.DiagonalUp,
            style.Borders.DiagonalDown,
            style.HorizontalAlignment.ToString(),
            style.VerticalAlignment.ToString(),
            style.WrapText,
            style.IndentLevel,
            style.TextRotation,
            style.ShrinkToFit,
            style.ReadingOrder,
            style.RelativeIndent,
            style.Number,
            style.Custom ?? string.Empty,
            style.IsLocked,
            style.IsHidden);
    }

    private static TResult ReadWithOpenXmlSdk<TResult>(string workbookPath, Func<OpenXmlReadContext, TResult> read)
    {
        SpreadsheetDocument? document = null;
        string? tempWorkbookPath = null;
        try
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), "Aspose.Cells_FOSS", "openxml-compare");
            Directory.CreateDirectory(tempDirectory);
            tempWorkbookPath = Path.Combine(tempDirectory, Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(workbookPath, tempWorkbookPath, true);

            document = SpreadsheetDocument.Open(tempWorkbookPath, false, new OpenSettings { AutoSave = false });
            var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
            var sharedStrings = LoadSharedStrings(workbookPart);
            var dateSystem = workbookPart.Workbook.WorkbookProperties?.Date1904?.Value == true ? DateSystem.Mac1904 : DateSystem.Windows1900;
            var styles = LoadOpenXmlStylesContext(workbookPart);
            return read(new OpenXmlReadContext(workbookPart, sharedStrings, dateSystem, styles));
        }
        finally
        {
            if (document is not null)
            {
                try
                {
                    document.Dispose();
                }
                catch (ObjectDisposedException)
                {
                    // Some malformed or edge-case packages trigger a second close inside the SDK on read-only dispose.
                }
            }

            if (!string.IsNullOrEmpty(tempWorkbookPath) && File.Exists(tempWorkbookPath))
            {
                try
                {
                    File.Delete(tempWorkbookPath);
                }
                catch (IOException)
                {
                    // Best effort cleanup for compare temp files.
                }
                catch (UnauthorizedAccessException)
                {
                    // Best effort cleanup for compare temp files.
                }
            }
        }
    }

    private static IEnumerable<OpenXmlCellDescriptor> EnumerateOpenXmlCells(OpenXmlReadContext context)
    {
        var sheets = context.WorkbookPart.Workbook.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>();
        foreach (var sheet in sheets)
        {
            if (sheet.Id?.Value is not string relationshipId)
            {
                continue;
            }

            if (context.WorkbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var sheetName = sheet.Name?.Value ?? "<unknown>";
            foreach (var cell in worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                if (!TryNormalizeCellReference(cell.CellReference?.Value, out var cellReference))
                {
                    continue;
                }

                yield return new OpenXmlCellDescriptor(sheetName, cellReference, cell);
            }
        }
    }

    private static CellSnapshot CreateOpenXmlCellSnapshot(string sheetName, string cellReference, DocumentFormat.OpenXml.Spreadsheet.Cell cell, IReadOnlyList<string> sharedStrings, ISet<int> dateStyleIndexes, DateSystem dateSystem)
    {
        var formula = NormalizeFormulaText(cell.CellFormula?.Text);
        var value = ReadOpenXmlValue(cell, sharedStrings, dateStyleIndexes, dateSystem, out var kind);
        var effectiveType = string.IsNullOrEmpty(formula) ? kind : "Formula";
        return new CellSnapshot(sheetName, cellReference, effectiveType, NormalizeValue(value), formula);
    }

    private static StyleSnapshot CreateOpenXmlStyleSnapshot(DocumentFormat.OpenXml.Spreadsheet.Cell cell, OpenXmlStylesContext styles)
    {
        var styleIndex = cell.StyleIndex?.Value is uint rawStyleIndex ? (int)rawStyleIndex : 0;
        if (styleIndex < 0 || styleIndex >= styles.CellFormats.Count)
        {
            styleIndex = 0;
        }

        return styles.CellFormats[styleIndex];
    }

    private static object? ReadOpenXmlValue(DocumentFormat.OpenXml.Spreadsheet.Cell cell, IReadOnlyList<string> sharedStrings, ISet<int> dateStyleIndexes, DateSystem dateSystem, out string kind)
    {
        kind = "Blank";
        var cellType = cell.DataType?.Value;
        var rawValue = cell.CellValue?.InnerText;
        var styleIndex = cell.StyleIndex?.Value;
        var isDateStyle = styleIndex.HasValue && dateStyleIndexes.Contains((int)styleIndex.Value);

        if (cellType == CellValues.InlineString)
        {
            kind = "String";
            return ReadInlineString(cell.InlineString);
        }

        if (cellType == CellValues.SharedString)
        {
            kind = "String";
            if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index >= 0 && index < sharedStrings.Count)
            {
                return sharedStrings[index];
            }

            return string.Empty;
        }

        if (cellType == CellValues.Boolean)
        {
            kind = "Boolean";
            return rawValue == "1" || string.Equals(rawValue, "true", StringComparison.OrdinalIgnoreCase);
        }

        if (cellType == CellValues.Date)
        {
            if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dateValue))
            {
                kind = "DateTime";
                return dateValue;
            }

            return null;
        }

        if (cellType == CellValues.String || cellType == CellValues.Error)
        {
            kind = "String";
            return rawValue ?? string.Empty;
        }

        if (string.IsNullOrEmpty(rawValue))
        {
            return null;
        }

        if (isDateStyle)
        {
            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
            {
                kind = "DateTime";
                return DateSerialConverter.FromSerial(serial, dateSystem == DateSystem.Mac1904 ? Aspose.Cells_FOSS.Core.DateSystem.Mac1904 : Aspose.Cells_FOSS.Core.DateSystem.Windows1900);
            }

            return null;
        }

        if (TryParseNumber(rawValue, out var numericValue, out kind))
        {
            return numericValue;
        }

        return null;
    }

    private static string ReadInlineString(InlineString? inlineString)
    {
        if (inlineString is null)
        {
            return string.Empty;
        }

        var texts = inlineString.Descendants<Text>().Select(delegate(Text text) { return text.Text; }).ToList();
        return texts.Count == 0 ? inlineString.InnerText : string.Concat(texts);
    }

    private static List<string> LoadSharedStrings(WorkbookPart workbookPart)
    {
        var table = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (table is null)
        {
            return new List<string>();
        }

        return table.Elements<SharedStringItem>()
            .Select(delegate(SharedStringItem item)
            {
                var texts = item.Descendants<Text>().Select(delegate(Text text) { return text.Text; }).ToList();
                return texts.Count == 0 ? item.InnerText : string.Concat(texts);
            })
            .ToList();
    }

    private static OpenXmlStylesContext LoadOpenXmlStylesContext(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart is null)
        {
            return OpenXmlStylesContext.Default;
        }

        using var stream = stylesPart.GetStream(FileMode.Open, FileAccess.Read);
        var document = XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        var root = document.Root;
        if (root is null)
        {
            return OpenXmlStylesContext.Default;
        }

        var customFormats = new Dictionary<int, string>();
        foreach (var numFmt in root.Element(SpreadsheetNs + "numFmts")?.Elements(SpreadsheetNs + "numFmt") ?? Enumerable.Empty<XElement>())
        {
            var id = ParseIntAttribute(numFmt.Attribute("numFmtId"));
            var formatCode = (string?)numFmt.Attribute("formatCode");
            if (id.HasValue && !string.IsNullOrEmpty(formatCode))
            {
                customFormats[id.Value] = formatCode;
            }
        }

        var fonts = (root.Element(SpreadsheetNs + "fonts")?.Elements(SpreadsheetNs + "font") ?? Enumerable.Empty<XElement>())
            .Select(ReadOpenXmlFont)
            .ToList();
        if (fonts.Count == 0)
        {
            fonts.Add(OpenXmlFontDescriptor.Default);
        }

        var fills = (root.Element(SpreadsheetNs + "fills")?.Elements(SpreadsheetNs + "fill") ?? Enumerable.Empty<XElement>())
            .Select(ReadOpenXmlFill)
            .ToList();
        if (fills.Count == 0)
        {
            fills.Add(OpenXmlFillDescriptor.Default);
        }

        var borders = (root.Element(SpreadsheetNs + "borders")?.Elements(SpreadsheetNs + "border") ?? Enumerable.Empty<XElement>())
            .Select(ReadOpenXmlBorder)
            .ToList();
        if (borders.Count == 0)
        {
            borders.Add(OpenXmlBorderDescriptor.Default);
        }

        var styleSnapshots = new List<StyleSnapshot>();
        var dateStyleIndexes = new HashSet<int>();
        var cellFormats = root.Element(SpreadsheetNs + "cellXfs")?.Elements(SpreadsheetNs + "xf").ToList() ?? new List<XElement>();
        if (cellFormats.Count == 0)
        {
            styleSnapshots.Add(StyleSnapshot.Default);
            return new OpenXmlStylesContext(styleSnapshots, dateStyleIndexes);
        }

        for (var index = 0; index < cellFormats.Count; index++)
        {
            var cellFormat = cellFormats[index];
            var fontId = ParseIntAttribute(cellFormat.Attribute("fontId")) ?? 0;
            var fillId = ParseIntAttribute(cellFormat.Attribute("fillId")) ?? 0;
            var borderId = ParseIntAttribute(cellFormat.Attribute("borderId")) ?? 0;
            var numberFormatId = ParseIntAttribute(cellFormat.Attribute("numFmtId")) ?? 0;
            var font = GetOrDefault(fonts, fontId, OpenXmlFontDescriptor.Default);
            var fill = GetOrDefault(fills, fillId, OpenXmlFillDescriptor.Default);
            var border = GetOrDefault(borders, borderId, OpenXmlBorderDescriptor.Default);
            var alignmentElement = cellFormat.Element(SpreadsheetNs + "alignment");
            var protectionElement = cellFormat.Element(SpreadsheetNs + "protection");
            var customFormatCode = customFormats.TryGetValue(numberFormatId, out var customFormat) ? customFormat : string.Empty;

            styleSnapshots.Add(new StyleSnapshot(
                string.Empty,
                string.Empty,
                font.Name,
                font.Size,
                font.Bold,
                font.Italic,
                font.Underline,
                font.StrikeThrough,
                font.Color,
                fill.Pattern,
                fill.ForegroundColor,
                fill.BackgroundColor,
                border.Left.Style,
                border.Left.Color,
                border.Right.Style,
                border.Right.Color,
                border.Top.Style,
                border.Top.Color,
                border.Bottom.Style,
                border.Bottom.Color,
                border.Diagonal.Style,
                border.Diagonal.Color,
                border.DiagonalUp,
                border.DiagonalDown,
                ParseHorizontalAlignment(alignmentElement?.Attribute("horizontal")?.Value),
                ParseVerticalAlignment(alignmentElement?.Attribute("vertical")?.Value),
                ParseBoolAttribute(alignmentElement?.Attribute("wrapText")),
                ParseIntAttribute(alignmentElement?.Attribute("indent")) ?? 0,
                ParseIntAttribute(alignmentElement?.Attribute("textRotation")) ?? 0,
                ParseBoolAttribute(alignmentElement?.Attribute("shrinkToFit")),
                ParseIntAttribute(alignmentElement?.Attribute("readingOrder")) ?? 0,
                ParseIntAttribute(alignmentElement?.Attribute("relativeIndent")) ?? 0,
                numberFormatId,
                customFormatCode,
                ParseOptionalBoolAttribute(protectionElement?.Attribute("locked")) ?? true,
                ParseBoolAttribute(protectionElement?.Attribute("hidden"))));

            if (BuiltInDateFormats.Contains(numberFormatId) || (!string.IsNullOrEmpty(customFormatCode) && LooksLikeDateFormat(customFormatCode)))
            {
                dateStyleIndexes.Add(index);
            }
        }

        return new OpenXmlStylesContext(styleSnapshots, dateStyleIndexes);
    }

    private static OpenXmlFontDescriptor ReadOpenXmlFont(XElement font)
    {
        return new OpenXmlFontDescriptor(
            (string?)font.Element(SpreadsheetNs + "name")?.Attribute("val") ?? "Calibri",
            NormalizeDouble(ParseDoubleValue(font.Element(SpreadsheetNs + "sz")?.Attribute("val")?.Value, 11d)),
            font.Element(SpreadsheetNs + "b") is not null,
            font.Element(SpreadsheetNs + "i") is not null,
            font.Element(SpreadsheetNs + "u") is not null,
            font.Element(SpreadsheetNs + "strike") is not null,
            NormalizeOpenXmlColor(font.Element(SpreadsheetNs + "color")));
    }

    private static OpenXmlFillDescriptor ReadOpenXmlFill(XElement fill)
    {
        var patternFill = fill.Element(SpreadsheetNs + "patternFill");
        if (patternFill is null)
        {
            return OpenXmlFillDescriptor.Default;
        }

        return new OpenXmlFillDescriptor(
            ParseFillPattern((string?)patternFill.Attribute("patternType")),
            NormalizeOpenXmlColor(patternFill.Element(SpreadsheetNs + "fgColor")),
            NormalizeOpenXmlColor(patternFill.Element(SpreadsheetNs + "bgColor")));
    }


    private static OpenXmlBorderDescriptor ReadOpenXmlBorder(XElement border)
    {
        return new OpenXmlBorderDescriptor(
            ReadOpenXmlBorderSide(border.Element(SpreadsheetNs + "left")),
            ReadOpenXmlBorderSide(border.Element(SpreadsheetNs + "right")),
            ReadOpenXmlBorderSide(border.Element(SpreadsheetNs + "top")),
            ReadOpenXmlBorderSide(border.Element(SpreadsheetNs + "bottom")),
            ReadOpenXmlBorderSide(border.Element(SpreadsheetNs + "diagonal")),
            ParseBoolAttribute(border.Attribute("diagonalUp")),
            ParseBoolAttribute(border.Attribute("diagonalDown")));
    }

    private static OpenXmlBorderSideDescriptor ReadOpenXmlBorderSide(XElement? side)
    {
        if (side is null)
        {
            return OpenXmlBorderSideDescriptor.Default;
        }

        return new OpenXmlBorderSideDescriptor(
            ParseBorderStyle(side.Attribute("style")?.Value),
            NormalizeOpenXmlColor(side.Element(SpreadsheetNs + "color")));
    }
}

