using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Testing;
using DocumentFormat.OpenXml.Packaging;

namespace Aspose.Cells_FOSS.DataValidationOpenXmlTests;

internal static class Program
{
    private static readonly XNamespace MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static int Main()
    {
        return TestRunner.Run(
            "DataValidation.OpenXmlTests",
            new TestCase("data_validations_match_openxml_sdk", DataValidationsMatchOpenXmlSdk),
            new TestCase("openxml_created_data_validations_load_in_library", OpenXmlCreatedDataValidationsLoadInLibrary));
    }

    private static void DataValidationsMatchOpenXmlSdk()
    {
        using var temp = new TemporaryDirectory("validation-openxml-write");
        var path = temp.GetPath("validations.xlsx");

        var workbook = ValidationScenarioFactory.CreateValidationWorkbook();
        workbook.Save(path);

        var loaded = new Workbook(path);
        ValidationScenarioFactory.AssertValidations(loaded);

        var snapshot = ReadValidationSnapshots(path, "Validation Sheet");
        AssertEx.Equal(3, snapshot.Count);

        AssertValidation(GetValidationSnapshot(snapshot, "A1:A3"), "A1:A3", "list", string.Empty, "stop", true, false, true, true, "Status", "Pick a status", "Invalid", "Choose from the list", "\"Open,Closed\"", string.Empty);
        AssertValidation(GetValidationSnapshot(snapshot, "B2:C3 E2:E3"), "B2:C3 E2:E3", "decimal", "between", "stop", false, true, false, true, string.Empty, string.Empty, "Range", "Enter 1.5-9.5", "1.5", "9.5");
        AssertValidation(GetValidationSnapshot(snapshot, "G1"), "G1", "custom", string.Empty, "warning", false, false, true, false, "Code", "Up to 5 chars", string.Empty, string.Empty, "LEN(G1)<=5", string.Empty);
    }

    private static void OpenXmlCreatedDataValidationsLoadInLibrary()
    {
        using var temp = new TemporaryDirectory("validation-openxml-read");
        var path = temp.GetPath("openxml-validations.xlsx");

        CreateOpenXmlValidationWorkbook(path);

        var loaded = new Workbook(path);
        var sheet = loaded.Worksheets[0];
        AssertEx.Equal("Validation Sheet", sheet.Name);
        AssertEx.Equal(2, sheet.Validations.Count);

        var listValidation = sheet.Validations[0];
        AssertEx.Equal(ValidationType.List, listValidation.Type);
        AssertEx.Equal("\"One,Two\"", listValidation.Formula1);
        AssertEx.True(listValidation.IgnoreBlank);
        AssertEx.True(listValidation.ShowInput);
        AssertEx.Equal("Status", listValidation.InputTitle);
        AssertEx.Equal("Pick one", listValidation.InputMessage);
        AssertEx.Equal(1, listValidation.Areas.Count);
        AssertEx.Equal(0, listValidation.Areas[0].FirstRow);
        AssertEx.Equal(0, listValidation.Areas[0].FirstColumn);
        AssertEx.Equal(2, listValidation.Areas[0].TotalRows);
        AssertEx.Equal(1, listValidation.Areas[0].TotalColumns);

        var decimalValidation = sheet.Validations[1];
        AssertEx.Equal(ValidationType.Decimal, decimalValidation.Type);
        AssertEx.Equal(OperatorType.Between, decimalValidation.Operator);
        AssertEx.Equal("1", decimalValidation.Formula1);
        AssertEx.Equal("9", decimalValidation.Formula2);
        AssertEx.False(decimalValidation.InCellDropDown);
        AssertEx.True(decimalValidation.ShowError);
        AssertEx.Equal("Range", decimalValidation.ErrorTitle);
        AssertEx.Equal("Enter 1-9", decimalValidation.ErrorMessage);
        AssertEx.Equal(2, decimalValidation.Areas.Count);
        AssertEx.Equal(1, decimalValidation.Areas[0].FirstRow);
        AssertEx.Equal(1, decimalValidation.Areas[0].FirstColumn);
        AssertEx.Equal(2, decimalValidation.Areas[0].TotalRows);
        AssertEx.Equal(2, decimalValidation.Areas[0].TotalColumns);
        AssertEx.Equal(1, decimalValidation.Areas[1].FirstRow);
        AssertEx.Equal(4, decimalValidation.Areas[1].FirstColumn);
        AssertEx.Equal(2, decimalValidation.Areas[1].TotalRows);
        AssertEx.Equal(1, decimalValidation.Areas[1].TotalColumns);

        var snapshot = ReadValidationSnapshots(path, "Validation Sheet");
        AssertEx.Equal(2, snapshot.Count);
        AssertValidation(GetValidationSnapshot(snapshot, "A1:A2"), "A1:A2", "list", string.Empty, "stop", true, false, true, false, "Status", "Pick one", string.Empty, string.Empty, "\"One,Two\"", string.Empty);
        AssertValidation(GetValidationSnapshot(snapshot, "B2:C3 E2:E3"), "B2:C3 E2:E3", "decimal", "between", "stop", false, true, false, true, string.Empty, string.Empty, "Range", "Enter 1-9", "1", "9");
    }

    private static List<ValidationSnapshot> ReadValidationSnapshots(string workbookPath, string sheetName)
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
        var validations = new List<ValidationSnapshot>();

        foreach (var validation in worksheetXml.Root?.Element(MainNs + "dataValidations")?.Elements(MainNs + "dataValidation") ?? Enumerable.Empty<XElement>())
        {
            validations.Add(new ValidationSnapshot(
                ((string?)validation.Attribute("sqref") ?? string.Empty).ToUpperInvariant(),
                (string?)validation.Attribute("type") ?? string.Empty,
                (string?)validation.Attribute("operator") ?? string.Empty,
                (string?)validation.Attribute("errorStyle") ?? "stop",
                ParseBool((string?)validation.Attribute("allowBlank")),
                ParseBool((string?)validation.Attribute("showDropDown")),
                ParseBool((string?)validation.Attribute("showInputMessage")),
                ParseBool((string?)validation.Attribute("showErrorMessage")),
                (string?)validation.Attribute("promptTitle") ?? string.Empty,
                (string?)validation.Attribute("prompt") ?? string.Empty,
                (string?)validation.Attribute("errorTitle") ?? string.Empty,
                (string?)validation.Attribute("error") ?? string.Empty,
                (string?)validation.Element(MainNs + "formula1") ?? string.Empty,
                (string?)validation.Element(MainNs + "formula2") ?? string.Empty));
        }

        return validations;
    }

    private static void CreateOpenXmlValidationWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
        WriteXmlPart(worksheetPart, "<?xml version=\"1.0\" encoding=\"utf-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:E3\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>One</t></is></c></row></sheetData><dataValidations count=\"2\"><dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" promptTitle=\"Status\" prompt=\"Pick one\" sqref=\"A1:A2\"><formula1>\"One,Two\"</formula1></dataValidation><dataValidation type=\"decimal\" operator=\"between\" showDropDown=\"1\" showErrorMessage=\"1\" errorTitle=\"Range\" error=\"Enter 1-9\" sqref=\"B2:C3 E2:E3\"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations></worksheet>");
        WriteXmlPart(workbookPart, "<?xml version=\"1.0\" encoding=\"utf-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Validation Sheet\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
    }

    private static void WriteXmlPart(OpenXmlPart part, string xml)
    {
        using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
        writer.Write(xml);
    }

    private static ValidationSnapshot GetValidationSnapshot(IReadOnlyList<ValidationSnapshot> snapshots, string sqref)
    {
        return snapshots.Single(delegate(ValidationSnapshot snapshot) { return string.Equals(snapshot.Sqref, sqref, StringComparison.Ordinal); });
    }

    private static void AssertValidation(ValidationSnapshot actual, string sqref, string type, string operatorName, string errorStyle, bool allowBlank, bool showDropDown, bool showInputMessage, bool showErrorMessage, string promptTitle, string prompt, string errorTitle, string error, string formula1, string formula2)
    {
        AssertEx.Equal(sqref, actual.Sqref);
        AssertEx.Equal(type, actual.Type);
        AssertEx.Equal(operatorName, actual.Operator);
        AssertEx.Equal(errorStyle, actual.ErrorStyle);
        AssertEx.Equal(allowBlank, actual.AllowBlank);
        AssertEx.Equal(showDropDown, actual.ShowDropDown);
        AssertEx.Equal(showInputMessage, actual.ShowInputMessage);
        AssertEx.Equal(showErrorMessage, actual.ShowErrorMessage);
        AssertEx.Equal(promptTitle, actual.PromptTitle);
        AssertEx.Equal(prompt, actual.Prompt);
        AssertEx.Equal(errorTitle, actual.ErrorTitle);
        AssertEx.Equal(error, actual.Error);
        AssertEx.Equal(formula1, actual.Formula1);
        AssertEx.Equal(formula2, actual.Formula2);
    }

    private static bool ParseBool(string? value)
    {
        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    private sealed record ValidationSnapshot(
        string Sqref,
        string Type,
        string Operator,
        string ErrorStyle,
        bool AllowBlank,
        bool ShowDropDown,
        bool ShowInputMessage,
        bool ShowErrorMessage,
        string PromptTitle,
        string Prompt,
        string ErrorTitle,
        string Error,
        string Formula1,
        string Formula2);
}

