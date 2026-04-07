using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Testing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aspose.Cells_FOSS.PageSetupOpenXmlTests;

internal static class Program
{
    private static readonly XNamespace MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static int Main()
    {
        return TestRunner.Run(
            "PageSetup.OpenXmlTests",
            new TestCase("page_setup_matches_openxml_sdk", PageSetupMatchesOpenXmlSdk),
            new TestCase("openxml_created_page_setup_loads_in_library", OpenXmlCreatedPageSetupLoadsInLibrary));
    }

    private static void PageSetupMatchesOpenXmlSdk()
    {
        using var temp = new TemporaryDirectory("page-setup-openxml-write");
        var path = temp.GetPath("page-setup.xlsx");

        var workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
        workbook.Save(path);

        var loaded = new Workbook(path);
        PageSetupScenarioFactory.AssertPageSetup(loaded);

        var snapshot = ReadPageSetupSnapshot(path, "Print Sheet");
        AssertEx.Equal("$A$1:$C$10", snapshot.PrintArea);
        AssertEx.Equal("$1:$2", snapshot.PrintTitleRows);
        AssertEx.Equal("$A:$B", snapshot.PrintTitleColumns);
        AssertEx.Equal(0.25d, snapshot.LeftMargin ?? 0d);
        AssertEx.Equal(0.4d, snapshot.RightMargin ?? 0d);
        AssertEx.Equal(0.5d, snapshot.TopMargin ?? 0d);
        AssertEx.Equal(0.6d, snapshot.BottomMargin ?? 0d);
        AssertEx.Equal(0.2d, snapshot.HeaderMargin ?? 0d);
        AssertEx.Equal(0.22d, snapshot.FooterMargin ?? 0d);
        AssertEx.Equal(9, snapshot.PaperSize ?? 0);
        AssertEx.Equal("landscape", snapshot.Orientation);
        AssertEx.Equal(3, snapshot.FirstPageNumber ?? 0);
        AssertEx.Equal(95, snapshot.Scale ?? 0);
        AssertEx.Equal(1, snapshot.FitToWidth ?? 0);
        AssertEx.Equal(2, snapshot.FitToHeight ?? 0);
        AssertEx.True(snapshot.PrintGridLines);
        AssertEx.True(snapshot.PrintHeadings);
        AssertEx.True(snapshot.HorizontalCentered);
        AssertEx.True(snapshot.VerticalCentered);
        AssertEx.Equal("&LLeft Header&CCenter Header&RRight Header", snapshot.OddHeader);
        AssertEx.Equal("&LLeft Footer&CCenter Footer&RRight Footer", snapshot.OddFooter);
        AssertEx.Equal(2, snapshot.RowBreaks.Count);
        AssertEx.Equal(4u, snapshot.RowBreaks[0]);
        AssertEx.Equal(7u, snapshot.RowBreaks[1]);
        AssertEx.Equal(1, snapshot.ColumnBreaks.Count);
        AssertEx.Equal(2u, snapshot.ColumnBreaks[0]);
    }

    private static void OpenXmlCreatedPageSetupLoadsInLibrary()
    {
        using var temp = new TemporaryDirectory("page-setup-openxml-read");
        var path = temp.GetPath("openxml-page-setup.xlsx");

        CreateOpenXmlPageSetupWorkbook(path);

        var loaded = new Workbook(path);
        PageSetupScenarioFactory.AssertPageSetup(loaded);

        var snapshot = ReadPageSetupSnapshot(path, "Print Sheet");
        AssertEx.Equal("$A$1:$C$10", snapshot.PrintArea);
        AssertEx.Equal("$1:$2", snapshot.PrintTitleRows);
        AssertEx.Equal("$A:$B", snapshot.PrintTitleColumns);
    }

    private static PageSetupSnapshot ReadPageSetupSnapshot(string workbookPath, string sheetName)
    {
        using var document = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
        var workbookXml = XDocument.Load(workbookPart.GetStream());
        var sheets = workbookXml.Root?.Element(MainNs + "sheets")?.Elements(MainNs + "sheet") ?? Enumerable.Empty<XElement>();
        var sheetElement = sheets.First(delegate(XElement element) { return string.Equals((string?)element.Attribute("name"), sheetName, StringComparison.Ordinal); });
        var sheetId = int.Parse((string?)sheetElement.Attribute("sheetId") ?? "1", CultureInfo.InvariantCulture) - 1;
        var relationshipId = (string?)sheetElement.Attribute(RelationshipNs + "id") ?? throw new InvalidOperationException("Worksheet relationship id is missing.");
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
        var worksheetXml = XDocument.Load(worksheetPart.GetStream());
        var root = worksheetXml.Root ?? throw new InvalidOperationException("Worksheet XML is empty.");

        var margins = root.Element(MainNs + "pageMargins");
        var pageSetup = root.Element(MainNs + "pageSetup");
        var printOptions = root.Element(MainNs + "printOptions");
        var headerFooter = root.Element(MainNs + "headerFooter");
        var rowBreaks = (root.Element(MainNs + "rowBreaks")?.Elements(MainNs + "brk") ?? Enumerable.Empty<XElement>())
            .Select(delegate(XElement element) { return uint.Parse((string?)element.Attribute("id") ?? "0", CultureInfo.InvariantCulture); })
            .ToList();
        var columnBreaks = (root.Element(MainNs + "colBreaks")?.Elements(MainNs + "brk") ?? Enumerable.Empty<XElement>())
            .Select(delegate(XElement element) { return uint.Parse((string?)element.Attribute("id") ?? "0", CultureInfo.InvariantCulture); })
            .ToList();

        string? printArea = null;
        string? printTitleRows = null;
        string? printTitleColumns = null;
        foreach (var definedName in workbookXml.Root?.Element(MainNs + "definedNames")?.Elements(MainNs + "definedName") ?? Enumerable.Empty<XElement>())
        {
            var localSheetId = (string?)definedName.Attribute("localSheetId");
            if (!string.Equals(localSheetId, sheetId.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal))
            {
                continue;
            }

            var name = (string?)definedName.Attribute("name") ?? string.Empty;
            if (string.Equals(name, "_xlnm.Print_Area", StringComparison.Ordinal))
            {
                printArea = UnqualifyDefinedNameValue(definedName.Value);
            }
            else if (string.Equals(name, "_xlnm.Print_Titles", StringComparison.Ordinal))
            {
                foreach (var segment in definedName.Value.Split(','))
                {
                    var unqualified = UnqualifyDefinedNameValue(segment);
                    if (unqualified.Contains('$') && unqualified.IndexOfAny("0123456789".ToCharArray()) >= 0 && unqualified.Contains(':') && char.IsDigit(unqualified.Replace("$", string.Empty)[0]))
                    {
                        printTitleRows = unqualified;
                    }
                    else
                    {
                        printTitleColumns = unqualified;
                    }
                }
            }
        }

        return new PageSetupSnapshot(
            printArea ?? string.Empty,
            printTitleRows ?? string.Empty,
            printTitleColumns ?? string.Empty,
            TryParseDouble((string?)margins?.Attribute("left")),
            TryParseDouble((string?)margins?.Attribute("right")),
            TryParseDouble((string?)margins?.Attribute("top")),
            TryParseDouble((string?)margins?.Attribute("bottom")),
            TryParseDouble((string?)margins?.Attribute("header")),
            TryParseDouble((string?)margins?.Attribute("footer")),
            TryParseInt((string?)pageSetup?.Attribute("paperSize")),
            (string?)pageSetup?.Attribute("orientation") ?? string.Empty,
            TryParseInt((string?)pageSetup?.Attribute("firstPageNumber")),
            TryParseInt((string?)pageSetup?.Attribute("scale")),
            TryParseInt((string?)pageSetup?.Attribute("fitToWidth")),
            TryParseInt((string?)pageSetup?.Attribute("fitToHeight")),
            ParseBool((string?)printOptions?.Attribute("gridLines")),
            ParseBool((string?)printOptions?.Attribute("headings")),
            ParseBool((string?)printOptions?.Attribute("horizontalCentered")),
            ParseBool((string?)printOptions?.Attribute("verticalCentered")),
            (string?)headerFooter?.Element(MainNs + "oddHeader") ?? string.Empty,
            (string?)headerFooter?.Element(MainNs + "oddFooter") ?? string.Empty,
            rowBreaks,
            columnBreaks);
    }

    private static void CreateOpenXmlPageSetupWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
        WriteXmlPart(worksheetPart, "<?xml version=\"1.0\" encoding=\"utf-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><dimension ref=\"A1:C10\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Title</t></is></c></row><row r=\"10\"><c r=\"C10\"><v>42</v></c></row></sheetData><printOptions headings=\"1\" gridLines=\"1\" gridLinesSet=\"1\" horizontalCentered=\"1\" verticalCentered=\"1\"/><pageMargins left=\"0.25\" right=\"0.4\" top=\"0.5\" bottom=\"0.6\" header=\"0.2\" footer=\"0.22\"/><pageSetup paperSize=\"9\" scale=\"95\" fitToWidth=\"1\" fitToHeight=\"2\" firstPageNumber=\"3\" useFirstPageNumber=\"1\" orientation=\"landscape\"/><headerFooter><oddHeader>&amp;LLeft Header&amp;CCenter Header&amp;RRight Header</oddHeader><oddFooter>&amp;LLeft Footer&amp;CCenter Footer&amp;RRight Footer</oddFooter></headerFooter><rowBreaks count=\"2\" manualBreakCount=\"2\"><brk id=\"4\" max=\"16383\" man=\"1\"/><brk id=\"7\" max=\"16383\" man=\"1\"/></rowBreaks><colBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"2\" max=\"1048575\" man=\"1\"/></colBreaks></worksheet>");
        WriteXmlPart(workbookPart, "<?xml version=\"1.0\" encoding=\"utf-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Print Sheet\" sheetId=\"1\" r:id=\"rId1\"/></sheets><definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">'Print Sheet'!$A$1:$C$10</definedName><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">'Print Sheet'!$1:$2,'Print Sheet'!$A:$B</definedName></definedNames></workbook>");
    }

    private static void WriteXmlPart(OpenXmlPart part, string xml)
    {
        using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
        writer.Write(xml);
    }

    private static string UnqualifyDefinedNameValue(string value)
    {
        var trimmed = value.Trim();
        var exclamationIndex = trimmed.LastIndexOf('!');
        return exclamationIndex >= 0 ? trimmed.Substring(exclamationIndex + 1) : trimmed;
    }

    private static int? TryParseInt(string? value)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : null;
    }

    private static double? TryParseDouble(string? value)
    {
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed : null;
    }

    private static bool ParseBool(string? value)
    {
        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    private sealed record PageSetupSnapshot(
        string PrintArea,
        string PrintTitleRows,
        string PrintTitleColumns,
        double? LeftMargin,
        double? RightMargin,
        double? TopMargin,
        double? BottomMargin,
        double? HeaderMargin,
        double? FooterMargin,
        int? PaperSize,
        string Orientation,
        int? FirstPageNumber,
        int? Scale,
        int? FitToWidth,
        int? FitToHeight,
        bool PrintGridLines,
        bool PrintHeadings,
        bool HorizontalCentered,
        bool VerticalCentered,
        string OddHeader,
        string OddFooter,
        IReadOnlyList<uint> RowBreaks,
        IReadOnlyList<uint> ColumnBreaks);
}
