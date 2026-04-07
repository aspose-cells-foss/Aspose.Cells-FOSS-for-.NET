using System.Xml.Linq;
using Aspose.Cells_FOSS.Testing;
using DocumentFormat.OpenXml.Packaging;

namespace Aspose.Cells_FOSS.ConditionalFormattingOpenXmlTests;

internal static class Program
{
    private static readonly XNamespace MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static int Main()
    {
        return TestRunner.Run(
            "ConditionalFormatting.OpenXmlTests",
            new TestCase("conditional_formattings_match_openxml_sdk", ConditionalFormattingsMatchOpenXmlSdk),
            new TestCase("advanced_conditional_formattings_match_openxml_sdk", AdvancedConditionalFormattingsMatchOpenXmlSdk),
            new TestCase("openxml_created_conditional_formattings_load_in_library", OpenXmlCreatedConditionalFormattingsLoadInLibrary));
    }

    private static void ConditionalFormattingsMatchOpenXmlSdk()
    {
        using var temp = new TemporaryDirectory("conditional-formatting-openxml-write");
        var path = temp.GetPath("conditional-formatting.xlsx");

        var workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
        workbook.Save(path);

        var loaded = new Workbook(path);
        ConditionalFormattingScenarioFactory.AssertConditionalFormattings(loaded);

        var snapshot = ReadSnapshots(path, "Conditional Formatting");
        AssertEx.Equal(2, snapshot.Count);
        AssertSnapshot(GetSnapshot(snapshot, "A1:A5"), "A1:A5", 2);
        AssertSnapshot(GetSnapshot(snapshot, "C1:C3"), "C1:C3", 1);
    }

    private static void AdvancedConditionalFormattingsMatchOpenXmlSdk()
    {
        using var temp = new TemporaryDirectory("conditional-formatting-openxml-advanced-write");
        var path = temp.GetPath("advanced-conditional-formatting.xlsx");

        var workbook = ConditionalFormattingScenarioFactory.CreateAdvancedConditionalFormattingWorkbook();
        workbook.Save(path);

        var loaded = new Workbook(path);
        ConditionalFormattingScenarioFactory.AssertAdvancedConditionalFormattings(loaded);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains("type=\"containsText\"", worksheetXml);
        AssertEx.Contains("text=\"error\"", worksheetXml);
        AssertEx.Contains("type=\"timePeriod\"", worksheetXml);
        AssertEx.Contains("timePeriod=\"today\"", worksheetXml);
        AssertEx.Contains("type=\"duplicateValues\"", worksheetXml);
        AssertEx.Contains("type=\"uniqueValues\"", worksheetXml);
        AssertEx.Contains("type=\"colorScale\"", worksheetXml);
        AssertEx.Contains("type=\"dataBar\"", worksheetXml);
        AssertEx.Contains("type=\"iconSet\"", worksheetXml);
    }
    private static void OpenXmlCreatedConditionalFormattingsLoadInLibrary()
    {
        using var temp = new TemporaryDirectory("conditional-formatting-openxml-read");
        var path = temp.GetPath("openxml-conditional-formatting.xlsx");

        CreateOpenXmlWorkbook(path);

        var loaded = new Workbook(path);
        AssertEx.Equal(1, loaded.Worksheets[0].ConditionalFormattings.Count);
        var collection = loaded.Worksheets[0].ConditionalFormattings[0];
        AssertEx.Equal(2, collection.Count);
        AssertEx.Equal(1, collection.RangeCount);
        AssertEx.Equal(FormatConditionType.CellValue, collection[0].Type);
        AssertEx.Equal(OperatorType.GreaterThan, collection[0].Operator);
        AssertEx.Equal("10", collection[0].Formula1);
        AssertEx.Equal(1, collection[0].Priority);
        AssertEx.True(collection[0].StopIfTrue);
        AssertEx.Equal(FillPattern.Solid, collection[0].Style.Pattern);
        AssertEx.Equal(Color.FromArgb(255, 255, 235, 156), collection[0].Style.ForegroundColor);
        AssertEx.Equal(FormatConditionType.Expression, collection[1].Type);
        AssertEx.Equal("A1=\"Y\"", collection[1].Formula1);
        AssertEx.Equal(2, collection[1].Priority);
    }

    private static List<ConditionalFormattingSnapshot> ReadSnapshots(string workbookPath, string sheetName)
    {
        using var document = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
        var workbookXml = XDocument.Load(workbookPart.GetStream());
        var sheetElement = workbookXml.Root?
            .Element(MainNs + "sheets")?
            .Elements(MainNs + "sheet")
            .First(delegate(XElement element) { return string.Equals((string?)element.Attribute("name"), sheetName, StringComparison.Ordinal); })
            ?? throw new InvalidOperationException("Worksheet was not found.");
        var relationshipId = (string?)sheetElement.Attribute(RelationshipNs + "id") ?? throw new InvalidOperationException("Worksheet relationship id is missing.");
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
        var worksheetXml = XDocument.Load(worksheetPart.GetStream());

        var snapshots = new List<ConditionalFormattingSnapshot>();
        foreach (var formatting in worksheetXml.Root?.Elements(MainNs + "conditionalFormatting") ?? Enumerable.Empty<XElement>())
        {
            var rules = new List<ConditionalFormattingRuleSnapshot>();
            foreach (var rule in formatting.Elements(MainNs + "cfRule"))
            {
                rules.Add(new ConditionalFormattingRuleSnapshot(
                    (string?)rule.Attribute("type") ?? string.Empty,
                    (string?)rule.Attribute("operator") ?? string.Empty,
                    (string?)rule.Attribute("dxfId") ?? string.Empty,
                    (string?)rule.Attribute("priority") ?? string.Empty,
                    (string?)rule.Attribute("stopIfTrue") ?? string.Empty,
                    rule.Elements(MainNs + "formula").Select(delegate(XElement element) { return element.Value; }).ToList()));
            }

            snapshots.Add(new ConditionalFormattingSnapshot(
                ((string?)formatting.Attribute("sqref") ?? string.Empty).ToUpperInvariant(),
                rules));
        }

        return snapshots;
    }

    private static ConditionalFormattingSnapshot GetSnapshot(IReadOnlyList<ConditionalFormattingSnapshot> snapshots, string sqref)
    {
        return snapshots.Single(delegate(ConditionalFormattingSnapshot snapshot) { return string.Equals(snapshot.Sqref, sqref, StringComparison.Ordinal); });
    }

    private static void AssertSnapshot(ConditionalFormattingSnapshot snapshot, string sqref, int ruleCount)
    {
        AssertEx.Equal(sqref, snapshot.Sqref);
        AssertEx.Equal(ruleCount, snapshot.Rules.Count);
    }

    private static void CreateOpenXmlWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
        var worksheetXml = new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            new XElement(MainNs + "worksheet",
                new XElement(MainNs + "dimension", new XAttribute("ref", "A1:A3")),
                new XElement(MainNs + "sheetData",
                    new XElement(MainNs + "row",
                        new XAttribute("r", 1),
                        new XElement(MainNs + "c",
                            new XAttribute("r", "A1"),
                            new XAttribute("t", "inlineStr"),
                            new XElement(MainNs + "is",
                                new XElement(MainNs + "t", "Y"))))),
                new XElement(MainNs + "conditionalFormatting",
                    new XAttribute("sqref", "A1:A3"),
                    new XElement(MainNs + "cfRule",
                        new XAttribute("type", "cellIs"),
                        new XAttribute("operator", "greaterThan"),
                        new XAttribute("dxfId", 0),
                        new XAttribute("priority", 1),
                        new XAttribute("stopIfTrue", 1),
                        new XElement(MainNs + "formula", "10")),
                    new XElement(MainNs + "cfRule",
                        new XAttribute("type", "expression"),
                        new XAttribute("priority", 2),
                        new XElement(MainNs + "formula", "A1=\"Y\"")))));
        WriteXmlPart(worksheetPart, worksheetXml.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
        WriteXmlPart(workbookPart, "<?xml version=\"1.0\" encoding=\"utf-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Conditional Formatting\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId2");
        WriteXmlPart(stylesPart, "<?xml version=\"1.0\" encoding=\"utf-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"1\"><dxf><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFEB9C\"/></patternFill></fill></dxf></dxfs></styleSheet>");
    }

    private static void WriteXmlPart(OpenXmlPart part, string xml)
    {
        using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
        writer.Write(xml);
    }

    private sealed record ConditionalFormattingSnapshot(string Sqref, IReadOnlyList<ConditionalFormattingRuleSnapshot> Rules);
    private sealed record ConditionalFormattingRuleSnapshot(string Type, string Operator, string DxfId, string Priority, string StopIfTrue, IReadOnlyList<string> Formulas);
}






