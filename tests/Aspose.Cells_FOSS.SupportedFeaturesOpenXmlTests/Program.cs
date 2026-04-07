using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.CompareOpenXml;
using Aspose.Cells_FOSS.Testing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static Aspose.Cells_FOSS.CompareOpenXml.ComparisonValueHelpers;
using static Aspose.Cells_FOSS.CompareOpenXml.OpenXmlComparisonSupport;

namespace Aspose.Cells_FOSS.SupportedFeaturesOpenXmlTests;

internal static class Program
{
    private static readonly XNamespace MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static int Main()
    {
        return TestRunner.Run(
            "SupportedFeatures.OpenXmlTests",
            new TestCase("supported_features_workbook_matches_openxml_snapshot", SupportedFeaturesWorkbookMatchesOpenXmlSnapshot),
            new TestCase("workbook_metadata_matches_openxml_snapshot", WorkbookMetadataMatchesOpenXmlSnapshot));
    }

    private static void SupportedFeaturesWorkbookMatchesOpenXmlSnapshot()
    {
        using var temp = new TemporaryDirectory("supported-features-openxml");
        var path = temp.GetPath("supported-features.xlsx");

        var workbook = CreateSupportedFeaturesWorkbook();
        workbook.Save(path, new SaveOptions { UseSharedStrings = true });

        var snapshot = ReadWorkbookSnapshot(path);

        AssertEx.True(snapshot.UsesDate1904);
        AssertEx.True(snapshot.HasSharedStrings);
        AssertEx.Equal(3, snapshot.Sheets.Count);
        AssertSheet(snapshot.Sheets[0], "All Features", "Visible");
        AssertSheet(snapshot.Sheets[1], "Target Sheet", "Hidden");
        AssertSheet(snapshot.Sheets[2], "VeryHidden Sheet", "VeryHidden");

        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "A1");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "B1");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "C1");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "D1");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "E1");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "F1");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "G1");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "H1");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "I2");
        AssertExpectedCell(workbook, snapshot.Cells, "All Features", "J3");
        AssertExpectedCell(workbook, snapshot.Cells, "Target Sheet", "C3");

        AssertExpectedStyle(workbook, snapshot.Styles, "All Features", "A1");
        AssertExpectedStyle(workbook, snapshot.Styles, "All Features", "C4");

        var worksheet = snapshot.Worksheets["All Features"];
        AssertEx.Equal("A1:J4", worksheet.Dimension);
        AssertRow(worksheet, 1, 22.5d, false);
        AssertRow(worksheet, 3, null, true);
        AssertColumn(worksheet, 0, 0, 18.25d, false);
        AssertColumn(worksheet, 2, 2, null, true);
        AssertEx.Equal(1, worksheet.MergeReferences.Count);
        AssertEx.Equal("A3:B4", worksheet.MergeReferences[0]);

        AssertEx.Equal(3, worksheet.Hyperlinks.Count);
        AssertHyperlink(worksheet.Hyperlinks[0], "H1", "https://example.com/docs?q=1", string.Empty, "External docs", "Docs");
        AssertHyperlink(worksheet.Hyperlinks[1], "I2", string.Empty, "'Target Sheet'!C3", "Jump to target", "Jump");
        AssertHyperlink(worksheet.Hyperlinks[2], "J3:K4", "mailto:test@example.com", string.Empty, "Send mail", "Mail");

        AssertEx.Equal(3, worksheet.Validations.Count);
        AssertValidation(GetValidationSnapshot(worksheet.Validations, "A1:A3"), "A1:A3", "list", string.Empty, "stop", true, false, true, true, "Status", "Pick a status", "Invalid", "Choose from the list", "\"Open,Closed\"", string.Empty);
        AssertValidation(GetValidationSnapshot(worksheet.Validations, "B2:C3 E2:E3"), "B2:C3 E2:E3", "decimal", "between", "stop", false, true, false, true, string.Empty, string.Empty, "Range", "Enter 1.5-9.5", "1.5", "9.5");
        AssertValidation(GetValidationSnapshot(worksheet.Validations, "G1"), "G1", "custom", string.Empty, "warning", false, false, true, false, "Code", "Up to 5 chars", string.Empty, string.Empty, "LEN(G1)<=5", string.Empty);

        AssertEx.Equal(2, worksheet.ConditionalFormattings.Count);
        AssertConditionalFormatting(GetConditionalFormattingSnapshot(worksheet.ConditionalFormattings, "A1:A5"), "A1:A5", 2);
        AssertConditionalFormatting(GetConditionalFormattingSnapshot(worksheet.ConditionalFormattings, "C1:C3"), "C1:C3", 1);

        var pageSetup = worksheet.PageSetup;
        AssertEx.Equal("$A$1:$J$10", pageSetup.PrintArea);
        AssertEx.Equal("$1:$2", pageSetup.PrintTitleRows);
        AssertEx.Equal("$A:$B", pageSetup.PrintTitleColumns);
        AssertEx.Equal(0.25d, pageSetup.LeftMargin ?? 0d);
        AssertEx.Equal(0.4d, pageSetup.RightMargin ?? 0d);
        AssertEx.Equal(0.5d, pageSetup.TopMargin ?? 0d);
        AssertEx.Equal(0.6d, pageSetup.BottomMargin ?? 0d);
        AssertEx.Equal(0.2d, pageSetup.HeaderMargin ?? 0d);
        AssertEx.Equal(0.22d, pageSetup.FooterMargin ?? 0d);
        AssertEx.Equal(9, pageSetup.PaperSize ?? 0);
        AssertEx.Equal("landscape", pageSetup.Orientation);
        AssertEx.Equal(3, pageSetup.FirstPageNumber ?? 0);
        AssertEx.Equal(95, pageSetup.Scale ?? 0);
        AssertEx.Equal(1, pageSetup.FitToWidth ?? 0);
        AssertEx.Equal(2, pageSetup.FitToHeight ?? 0);
        AssertEx.True(pageSetup.PrintGridLines);
        AssertEx.True(pageSetup.PrintHeadings);
        AssertEx.True(pageSetup.HorizontalCentered);
        AssertEx.True(pageSetup.VerticalCentered);
        AssertEx.Equal("&LLeft Header&CCenter Header&RRight Header", pageSetup.OddHeader);
        AssertEx.Equal("&LLeft Footer&CCenter Footer&RRight Footer", pageSetup.OddFooter);
        AssertEx.Equal(2, pageSetup.RowBreaks.Count);
        AssertEx.Equal(4u, pageSetup.RowBreaks[0]);
        AssertEx.Equal(7u, pageSetup.RowBreaks[1]);
        AssertEx.Equal(1, pageSetup.ColumnBreaks.Count);
        AssertEx.Equal(2u, pageSetup.ColumnBreaks[0]);
        var definedNames = ReadDefinedNames(path);
        AssertEx.Equal(2, definedNames.Count);
        AssertDefinedName(definedNames[0], "GlobalAmount", "'All Features'!$A$1:$A$3", null, true, "Primary selection");
        AssertDefinedName(definedNames[1], "LocalTarget", "'Target Sheet'!$C$3", 1, false, string.Empty);
    }


    private static void WorkbookMetadataMatchesOpenXmlSnapshot()
    {
        using var temp = new TemporaryDirectory("workbook-metadata-openxml");
        var path = temp.GetPath("workbook-metadata.xlsx");

        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        workbook.Save(path);

        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var workbookPart = document.WorkbookPart!;
            var workbookRoot = workbookPart.Workbook;
            AssertEx.Equal("WorkbookCode", workbookRoot.WorkbookProperties?.CodeName?.Value);
            AssertEx.True(workbookRoot.WorkbookProperties?.ShowObjects is not null);
            AssertEx.True(workbookRoot.WorkbookProtection is not null);
            AssertEx.Equal<uint>(1u, workbookRoot.BookViews?.Elements<DocumentFormat.OpenXml.Spreadsheet.WorkbookView>().First().ActiveTab?.Value ?? 0u);
            AssertEx.Equal<uint>(1u, workbookRoot.BookViews?.Elements<DocumentFormat.OpenXml.Spreadsheet.WorkbookView>().First().FirstSheet?.Value ?? 0u);
            AssertEx.True(workbookRoot.CalculationProperties?.CalculationMode is not null);
            AssertEx.True(workbookRoot.CalculationProperties?.ReferenceMode is not null);
        }

        var workbookXml = ZipPackageHelper.ReadEntryText(path, "xl/workbook.xml");
        var coreXml = ZipPackageHelper.ReadEntryText(path, "docProps/core.xml");
        var appXml = ZipPackageHelper.ReadEntryText(path, "docProps/app.xml");
        AssertEx.Contains("showObjects=\"placeholders\"", workbookXml);
        AssertEx.Contains("calcMode=\"manual\"", workbookXml);
        AssertEx.Contains("refMode=\"R1C1\"", workbookXml);
        AssertEx.Contains("Quarterly Summary", coreXml);
        AssertEx.Contains("Automation", coreXml);
        AssertEx.Contains("Aspose.Cells_FOSS Tests", appXml);
        AssertEx.Contains("https://example.com/base/", appXml);

        var loaded = new Workbook(path);
        WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(loaded);
    }

    private static Workbook CreateSupportedFeaturesWorkbook()
    {
        var workbook = new Workbook();
        workbook.Settings.Date1904 = true;

        var sheet = workbook.Worksheets[0];
        sheet.Name = "All Features";

        var targetIndex = workbook.Worksheets.Add("Target Sheet");
        var targetSheet = workbook.Worksheets[targetIndex];
        targetSheet.VisibilityType = VisibilityType.Hidden;
        targetSheet.Cells[2, 2].PutValue("Target");

        var veryHiddenIndex = workbook.Worksheets.Add("VeryHidden Sheet");
        workbook.Worksheets[veryHiddenIndex].VisibilityType = VisibilityType.VeryHidden;

        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["C1"].PutValue(true);
        sheet.Cells["D1"].PutValue(12.5m);
        sheet.Cells["E1"].PutValue(6.02214076E+23);
        sheet.Cells["F1"].PutValue(new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc));
        sheet.Cells["G1"].PutValue(20);
        sheet.Cells["G1"].Formula = "=B1*2";
        sheet.Cells["A3"].PutValue("Merged");
        sheet.Cells["C4"].PutValue(42.1234m);

        var primaryStyle = sheet.Cells["A1"].GetStyle();
        StyleScenarioFactory.ApplyPrimaryStyle(primaryStyle);
        sheet.Cells["A1"].SetStyle(primaryStyle);

        var customStyle = sheet.Cells["C4"].GetStyle();
        StyleScenarioFactory.ApplyCustomNumberStyle(customStyle);
        sheet.Cells["C4"].SetStyle(customStyle);

        sheet.Cells.Rows[1].Height = 22.5d;
        sheet.Cells.Rows[3].IsHidden = true;
        sheet.Cells.Columns[0].Width = 18.25d;
        sheet.Cells.Columns[2].IsHidden = true;
        sheet.Cells.Merge(2, 0, 2, 2);

        sheet.Cells["H1"].PutValue("Docs");
        var external = sheet.Hyperlinks[sheet.Hyperlinks.Add("H1", 1, 1, "https://example.com/docs?q=1")];
        external.TextToDisplay = "Docs";
        external.ScreenTip = "External docs";

        sheet.Cells["I2"].PutValue("Jump");
        var internalLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("I2", 1, 1, "'Target Sheet'!C3")];
        internalLink.TextToDisplay = "Jump";
        internalLink.ScreenTip = "Jump to target";

        sheet.Cells["J3"].PutValue("Mail");
        var rangeLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("J3", 2, 2, "mailto:test@example.com")];
        rangeLink.TextToDisplay = "Mail";
        rangeLink.ScreenTip = "Send mail";

        var listValidation = sheet.Validations[sheet.Validations.Add(CellArea.CreateCellArea("A1", "A3"))];
        listValidation.Type = ValidationType.List;
        listValidation.Formula1 = "\"Open,Closed\"";
        listValidation.IgnoreBlank = true;
        listValidation.ShowInput = true;
        listValidation.InputTitle = "Status";
        listValidation.InputMessage = "Pick a status";
        listValidation.ShowError = true;
        listValidation.ErrorTitle = "Invalid";
        listValidation.ErrorMessage = "Choose from the list";

        var decimalValidation = sheet.Validations[sheet.Validations.Add(CellArea.CreateCellArea("B2", "C3"))];
        decimalValidation.Type = ValidationType.Decimal;
        decimalValidation.Operator = OperatorType.Between;
        decimalValidation.Formula1 = "1.5";
        decimalValidation.Formula2 = "9.5";
        decimalValidation.InCellDropDown = false;
        decimalValidation.ShowError = true;
        decimalValidation.ErrorTitle = "Range";
        decimalValidation.ErrorMessage = "Enter 1.5-9.5";
        decimalValidation.AddArea(CellArea.CreateCellArea("E2", "E3"));

        var customValidation = sheet.Validations[sheet.Validations.Add(CellArea.CreateCellArea("G1", "G1"))];
        customValidation.Type = ValidationType.Custom;
        customValidation.AlertStyle = ValidationAlertType.Warning;
        customValidation.Formula1 = "LEN(G1)<=5";
        customValidation.ShowInput = true;
        customValidation.InputTitle = "Code";
        customValidation.InputMessage = "Up to 5 chars";

        var primaryConditionalFormatting = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        primaryConditionalFormatting.AddArea(CellArea.CreateCellArea("A1", "A5"));
        var betweenCondition = primaryConditionalFormatting[primaryConditionalFormatting.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "5", "15")];
        betweenCondition.StopIfTrue = true;
        var betweenStyle = betweenCondition.Style;
        betweenStyle.Pattern = FillPattern.Solid;
        betweenStyle.ForegroundColor = Color.FromArgb(255, 255, 199, 206);
        betweenStyle.Font.Bold = true;
        betweenStyle.Font.Color = Color.FromArgb(255, 156, 0, 6);
        betweenCondition.Style = betweenStyle;
        var expressionCondition = primaryConditionalFormatting[primaryConditionalFormatting.AddCondition(FormatConditionType.Expression, OperatorType.None, "MOD(A1,2)=0", string.Empty)];
        var expressionStyle = expressionCondition.Style;
        expressionStyle.Font.Italic = true;
        expressionStyle.Font.Color = Color.FromArgb(255, 0, 0, 255);
        expressionCondition.Style = expressionStyle;

        var secondaryConditionalFormatting = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        secondaryConditionalFormatting.AddArea(CellArea.CreateCellArea("C1", "C3"));
        var greaterThanCondition = secondaryConditionalFormatting[secondaryConditionalFormatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", string.Empty)];
        var greaterThanStyle = greaterThanCondition.Style;
        greaterThanStyle.Pattern = FillPattern.Solid;
        greaterThanStyle.ForegroundColor = Color.FromArgb(255, 198, 239, 206);
        greaterThanStyle.Font.Color = Color.FromArgb(255, 0, 97, 0);
        greaterThanCondition.Style = greaterThanStyle;

        var pageSetup = sheet.PageSetup;
        pageSetup.LeftMarginInch = 0.25d;
        pageSetup.RightMarginInch = 0.4d;
        pageSetup.TopMarginInch = 0.5d;
        pageSetup.BottomMarginInch = 0.6d;
        pageSetup.HeaderMarginInch = 0.2d;
        pageSetup.FooterMarginInch = 0.22d;
        pageSetup.Orientation = PageOrientationType.Landscape;
        pageSetup.PaperSize = PaperSizeType.PaperA4;
        pageSetup.FirstPageNumber = 3;
        pageSetup.Scale = 95;
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = 2;
        pageSetup.PrintArea = "$A$1:$J$10";
        pageSetup.PrintTitleRows = "$1:$2";
        pageSetup.PrintTitleColumns = "$A:$B";
        pageSetup.LeftHeader = "Left Header";
        pageSetup.CenterHeader = "Center Header";
        pageSetup.RightHeader = "Right Header";
        pageSetup.LeftFooter = "Left Footer";
        pageSetup.CenterFooter = "Center Footer";
        pageSetup.RightFooter = "Right Footer";
        pageSetup.PrintGridlines = true;
        pageSetup.PrintHeadings = true;
        pageSetup.CenterHorizontally = true;
        pageSetup.CenterVertically = true;
        pageSetup.AddHorizontalPageBreak(4);
        pageSetup.AddHorizontalPageBreak(7);
        pageSetup.AddVerticalPageBreak(2);

        var globalDefinedName = workbook.DefinedNames[workbook.DefinedNames.Add("GlobalAmount", "'All Features'!$A$1:$A$3")];
        globalDefinedName.Hidden = true;
        globalDefinedName.Comment = "Primary selection";
        workbook.DefinedNames.Add("LocalTarget", "'Target Sheet'!$C$3", 1);

        return workbook;
    }

    private static void AssertExpectedCell(Workbook workbook, IReadOnlyDictionary<string, CellSnapshot> actualCells, string sheetName, string cellReference)
    {
        var sheet = workbook.Worksheets[sheetName];
        var cell = sheet.Cells[cellReference];
        var expected = new CellSnapshot(
            sheetName,
            cellReference,
            NormalizeCellType(cell.Type),
            NormalizeValue(cell.Value),
            NormalizeFormulaText(cell.Formula));
        var key = BuildCellKey(sheetName, cellReference);
        AssertEx.True(actualCells.TryGetValue(key, out var actual), $"Expected Open XML cell snapshot for {key}.");

        AssertEx.Equal(expected.SheetName, actual!.SheetName);
        AssertEx.Equal(expected.CellReference, actual.CellReference);
        AssertEx.Equal(expected.CellType, actual.CellType);
        AssertEx.Equal(expected.Formula, actual.Formula);

        if (cell.Value is DateTime expectedDateTime)
        {
            var actualDateTime = DateTime.Parse(actual.Value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
            AssertEx.Equal(expectedDateTime.Ticks, actualDateTime.Ticks);
            return;
        }

        AssertEx.Equal(expected.Value, actual.Value);
    }

    private static void AssertExpectedStyle(Workbook workbook, IReadOnlyDictionary<string, StyleSnapshot> actualStyles, string sheetName, string cellReference)
    {
        var expected = CreateLibraryStyleSnapshot(sheetName, cellReference, workbook.Worksheets[sheetName].Cells[cellReference].GetStyle());
        var key = BuildCellKey(sheetName, cellReference);
        AssertEx.True(actualStyles.TryGetValue(key, out var actual), $"Expected Open XML style snapshot for {key}.");

        AssertEx.Equal(expected.SheetName, actual!.SheetName);
        AssertEx.Equal(expected.CellReference, actual.CellReference);
        AssertEx.Equal(expected.FontName, actual.FontName);
        AssertEx.Equal(expected.FontSize, actual.FontSize);
        AssertEx.Equal(expected.FontBold, actual.FontBold);
        AssertEx.Equal(expected.FontItalic, actual.FontItalic);
        AssertEx.Equal(expected.FontUnderline, actual.FontUnderline);
        AssertEx.Equal(expected.FontStrikeThrough, actual.FontStrikeThrough);
        AssertEx.Equal(expected.FontColor, actual.FontColor);
        AssertEx.Equal(expected.FillPattern, actual.FillPattern);
        AssertEx.Equal(expected.FillForegroundColor, actual.FillForegroundColor);
        AssertEx.Equal(expected.FillBackgroundColor, actual.FillBackgroundColor);
        AssertEx.Equal(expected.LeftBorderStyle, actual.LeftBorderStyle);
        AssertEx.Equal(expected.LeftBorderColor, actual.LeftBorderColor);
        AssertEx.Equal(expected.RightBorderStyle, actual.RightBorderStyle);
        AssertEx.Equal(expected.RightBorderColor, actual.RightBorderColor);
        AssertEx.Equal(expected.TopBorderStyle, actual.TopBorderStyle);
        AssertEx.Equal(expected.TopBorderColor, actual.TopBorderColor);
        AssertEx.Equal(expected.BottomBorderStyle, actual.BottomBorderStyle);
        AssertEx.Equal(expected.BottomBorderColor, actual.BottomBorderColor);
        AssertEx.Equal(expected.DiagonalBorderStyle, actual.DiagonalBorderStyle);
        AssertEx.Equal(expected.DiagonalBorderColor, actual.DiagonalBorderColor);
        AssertEx.Equal(expected.DiagonalUp, actual.DiagonalUp);
        AssertEx.Equal(expected.DiagonalDown, actual.DiagonalDown);
        AssertEx.Equal(expected.HorizontalAlignment, actual.HorizontalAlignment);
        AssertEx.Equal(expected.VerticalAlignment, actual.VerticalAlignment);
        AssertEx.Equal(expected.WrapText, actual.WrapText);
        AssertEx.Equal(expected.IndentLevel, actual.IndentLevel);
        AssertEx.Equal(expected.TextRotation, actual.TextRotation);
        AssertEx.Equal(expected.ShrinkToFit, actual.ShrinkToFit);
        AssertEx.Equal(expected.ReadingOrder, actual.ReadingOrder);
        AssertEx.Equal(expected.RelativeIndent, actual.RelativeIndent);
        AssertEx.Equal(expected.IsLocked, actual.IsLocked);
        AssertEx.Equal(expected.IsHidden, actual.IsHidden);

        if (!string.IsNullOrEmpty(expected.NumberFormatCode))
        {
            AssertEx.Equal(expected.NumberFormatCode, actual.NumberFormatCode);
        }
        else
        {
            AssertEx.Equal(expected.NumberFormatId, actual.NumberFormatId);
        }
    }
    private static string NormalizeCellType(CellValueType type)
    {
        switch (type)
        {
            case CellValueType.String:
                return "String";
            case CellValueType.Number:
                return "Number";
            case CellValueType.Boolean:
                return "Boolean";
            case CellValueType.DateTime:
                return "DateTime";
            case CellValueType.Formula:
                return "Formula";
            default:
                return "Blank";
        }
    }

    private static WorkbookSnapshot ReadWorkbookSnapshot(string workbookPath)
    {
        var cells = ReadCellDataWithOpenXmlSdk(workbookPath);
        var styles = ReadCellStylesWithOpenXmlSdk(workbookPath);

        using var document = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
        var workbookXml = XDocument.Load(workbookPart.GetStream());
        var workbookRoot = workbookXml.Root ?? throw new InvalidOperationException("Workbook XML is empty.");
        var sheets = new List<SheetSnapshot>();
        var worksheetSnapshots = new Dictionary<string, WorksheetSnapshot>(StringComparer.Ordinal);

        foreach (var sheetElement in workbookRoot.Element(MainNs + "sheets")?.Elements(MainNs + "sheet") ?? Enumerable.Empty<XElement>())
        {
            var relationshipId = (string?)sheetElement.Attribute(RelationshipNs + "id");
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            if (workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var sheetName = (string?)sheetElement.Attribute("name") ?? string.Empty;
            var visibility = ResolveVisibility((string?)sheetElement.Attribute("state"));
            sheets.Add(new SheetSnapshot(sheetName, visibility));
            worksheetSnapshots[sheetName] = ReadWorksheetSnapshot(workbookXml, workbookPart, sheetElement, worksheetPart);
        }

        var usesDate1904 = workbookPart.Workbook.WorkbookProperties?.Date1904?.Value == true;
        var hasSharedStrings = workbookPart.SharedStringTablePart is not null;
        return new WorkbookSnapshot(usesDate1904, hasSharedStrings, sheets, cells, styles, worksheetSnapshots);
    }

    private static WorksheetSnapshot ReadWorksheetSnapshot(XDocument workbookXml, WorkbookPart workbookPart, XElement sheetElement, WorksheetPart worksheetPart)
    {
        var worksheetXml = XDocument.Load(worksheetPart.GetStream());
        var worksheetRoot = worksheetXml.Root ?? throw new InvalidOperationException("Worksheet XML is empty.");
        var sheetName = (string?)sheetElement.Attribute("name") ?? string.Empty;
        var sheetIndex = int.Parse((string?)sheetElement.Attribute("sheetId") ?? "1", CultureInfo.InvariantCulture) - 1;

        var rows = new Dictionary<int, RowSnapshot>();
        foreach (var row in worksheetRoot.Element(MainNs + "sheetData")?.Elements(MainNs + "row") ?? Enumerable.Empty<XElement>())
        {
            if (!int.TryParse((string?)row.Attribute("r"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex))
            {
                continue;
            }

            rows[rowIndex - 1] = new RowSnapshot(TryParseDouble((string?)row.Attribute("ht")), ParseBool((string?)row.Attribute("hidden")));
        }

        var columns = new List<ColumnSnapshot>();
        foreach (var column in worksheetRoot.Element(MainNs + "cols")?.Elements(MainNs + "col") ?? Enumerable.Empty<XElement>())
        {
            if (!int.TryParse((string?)column.Attribute("min"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var minColumnIndex)
                || !int.TryParse((string?)column.Attribute("max"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var maxColumnIndex))
            {
                continue;
            }

            columns.Add(new ColumnSnapshot(minColumnIndex - 1, maxColumnIndex - 1, TryParseDouble((string?)column.Attribute("width")), ParseBool((string?)column.Attribute("hidden"))));
        }

        var merges = new List<string>();
        foreach (var mergeCell in worksheetRoot.Element(MainNs + "mergeCells")?.Elements(MainNs + "mergeCell") ?? Enumerable.Empty<XElement>())
        {
            var reference = ((string?)mergeCell.Attribute("ref") ?? string.Empty).ToUpperInvariant();
            if (!string.IsNullOrEmpty(reference))
            {
                merges.Add(reference);
            }
        }

        var hyperlinkTargets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var relationship in worksheetPart.HyperlinkRelationships)
        {
            hyperlinkTargets[relationship.Id] = relationship.Uri.ToString();
        }

        var hyperlinks = new List<HyperlinkSnapshot>();
        foreach (var hyperlink in worksheetRoot.Element(MainNs + "hyperlinks")?.Elements(MainNs + "hyperlink") ?? Enumerable.Empty<XElement>())
        {
            var relationshipId = (string?)hyperlink.Attribute(RelationshipNs + "id");
            hyperlinks.Add(new HyperlinkSnapshot(
                ((string?)hyperlink.Attribute("ref") ?? string.Empty).ToUpperInvariant(),
                relationshipId is not null && hyperlinkTargets.TryGetValue(relationshipId, out var address) ? address : string.Empty,
                (string?)hyperlink.Attribute("location") ?? string.Empty,
                (string?)hyperlink.Attribute("tooltip") ?? string.Empty,
                (string?)hyperlink.Attribute("display") ?? string.Empty));
        }

        var validations = new List<ValidationSnapshot>();
        foreach (var validation in worksheetRoot.Element(MainNs + "dataValidations")?.Elements(MainNs + "dataValidation") ?? Enumerable.Empty<XElement>())
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

        var conditionalFormattings = new List<ConditionalFormattingSnapshot>();
        foreach (var formatting in worksheetRoot.Elements(MainNs + "conditionalFormatting"))
        {
            var rules = new List<ConditionalFormattingRuleSnapshot>();
            foreach (var rule in formatting.Elements(MainNs + "cfRule"))
            {
                rules.Add(new ConditionalFormattingRuleSnapshot(
                    (string?)rule.Attribute("type") ?? string.Empty,
                    (string?)rule.Attribute("operator") ?? string.Empty,
                    (string?)rule.Attribute("priority") ?? string.Empty,
                    (string?)rule.Attribute("stopIfTrue") ?? string.Empty,
                    (string?)rule.Attribute("dxfId") ?? string.Empty,
                    rule.Elements(MainNs + "formula").Select(delegate(XElement element) { return element.Value; }).ToList()));
            }

            conditionalFormattings.Add(new ConditionalFormattingSnapshot(
                ((string?)formatting.Attribute("sqref") ?? string.Empty).ToUpperInvariant(),
                rules));
        }

        var pageSetup = ReadPageSetupSnapshot(workbookXml, workbookPart, worksheetRoot, sheetName, sheetIndex);
        return new WorksheetSnapshot(
            sheetName,
            ResolveVisibility((string?)sheetElement.Attribute("state")),
            ((string?)worksheetRoot.Element(MainNs + "dimension")?.Attribute("ref") ?? string.Empty).ToUpperInvariant(),
            rows,
            columns,
            merges,
            hyperlinks,
            validations,
            conditionalFormattings,
            pageSetup);
    }

    private static PageSetupSnapshot ReadPageSetupSnapshot(XDocument workbookXml, WorkbookPart workbookPart, XElement worksheetRoot, string sheetName, int sheetIndex)
    {
        var margins = worksheetRoot.Element(MainNs + "pageMargins");
        var pageSetup = worksheetRoot.Element(MainNs + "pageSetup");
        var printOptions = worksheetRoot.Element(MainNs + "printOptions");
        var headerFooter = worksheetRoot.Element(MainNs + "headerFooter");
        var rowBreaks = (worksheetRoot.Element(MainNs + "rowBreaks")?.Elements(MainNs + "brk") ?? Enumerable.Empty<XElement>())
            .Select(delegate(XElement element) { return uint.Parse((string?)element.Attribute("id") ?? "0", CultureInfo.InvariantCulture); })
            .ToList();
        var columnBreaks = (worksheetRoot.Element(MainNs + "colBreaks")?.Elements(MainNs + "brk") ?? Enumerable.Empty<XElement>())
            .Select(delegate(XElement element) { return uint.Parse((string?)element.Attribute("id") ?? "0", CultureInfo.InvariantCulture); })
            .ToList();

        string? printArea = null;
        string? printTitleRows = null;
        string? printTitleColumns = null;
        foreach (var definedName in workbookXml.Root?.Element(MainNs + "definedNames")?.Elements(MainNs + "definedName") ?? Enumerable.Empty<XElement>())
        {
            var localSheetId = (string?)definedName.Attribute("localSheetId");
            if (!string.Equals(localSheetId, sheetIndex.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal))
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

    private static IReadOnlyList<DefinedNameSnapshot> ReadDefinedNames(string workbookPath)
    {
        using var document = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
        var workbookXml = XDocument.Load(workbookPart.GetStream());
        var definedNames = new List<DefinedNameSnapshot>();

        foreach (var definedName in workbookXml.Root?.Element(MainNs + "definedNames")?.Elements(MainNs + "definedName") ?? Enumerable.Empty<XElement>())
        {
            var name = (string?)definedName.Attribute("name") ?? string.Empty;
            if (string.Equals(name, "_xlnm.Print_Area", StringComparison.Ordinal)
                || string.Equals(name, "_xlnm.Print_Titles", StringComparison.Ordinal))
            {
                continue;
            }

            definedNames.Add(new DefinedNameSnapshot(
                name,
                definedName.Value,
                TryParseInt((string?)definedName.Attribute("localSheetId")),
                ParseBool((string?)definedName.Attribute("hidden")),
                (string?)definedName.Attribute("comment") ?? string.Empty));
        }

        return definedNames;
    }

    private static void AssertDefinedName(DefinedNameSnapshot actual, string name, string formula, int? localSheetIndex, bool hidden, string comment)
    {
        AssertEx.Equal(name, actual.Name);
        AssertEx.Equal(formula, actual.Formula);
        AssertEx.Equal(localSheetIndex, actual.LocalSheetIndex);
        AssertEx.Equal(hidden, actual.Hidden);
        AssertEx.Equal(comment, actual.Comment);
    }

    private static void AssertSheet(SheetSnapshot actual, string expectedName, string expectedVisibility)
    {
        AssertEx.Equal(expectedName, actual.Name);
        AssertEx.Equal(expectedVisibility, actual.Visibility);
    }

    private static void AssertRow(WorksheetSnapshot snapshot, int rowIndex, double? height, bool hidden)
    {
        AssertEx.True(snapshot.Rows.TryGetValue(rowIndex, out var row), $"Expected row {rowIndex} to exist in Open XML snapshot.");
        AssertEx.Equal(height, row!.Height);
        AssertEx.Equal(hidden, row.Hidden);
    }

    private static void AssertColumn(WorksheetSnapshot snapshot, int minColumnIndex, int maxColumnIndex, double? width, bool hidden)
    {
        var column = snapshot.Columns.SingleOrDefault(delegate(ColumnSnapshot item) { return item.MinColumnIndex == minColumnIndex && item.MaxColumnIndex == maxColumnIndex; });
        AssertEx.NotNull(column, $"Expected column range {minColumnIndex}:{maxColumnIndex} to exist in Open XML snapshot.");
        AssertEx.Equal(width, column!.Width);
        AssertEx.Equal(hidden, column.Hidden);
    }

    private static void AssertHyperlink(HyperlinkSnapshot actual, string area, string address, string location, string screenTip, string display)
    {
        AssertEx.Equal(area, actual.Area);
        AssertEx.Equal(address, actual.Address);
        AssertEx.Equal(location, actual.Location);
        AssertEx.Equal(screenTip, actual.ScreenTip);
        AssertEx.Equal(display, actual.Display);
    }

    private static ValidationSnapshot GetValidationSnapshot(IReadOnlyList<ValidationSnapshot> snapshots, string sqref)
    {
        return snapshots.Single(delegate(ValidationSnapshot snapshot) { return string.Equals(snapshot.Sqref, sqref, StringComparison.Ordinal); });
    }

    private static ConditionalFormattingSnapshot GetConditionalFormattingSnapshot(IReadOnlyList<ConditionalFormattingSnapshot> snapshots, string sqref)
    {
        return snapshots.Single(delegate(ConditionalFormattingSnapshot snapshot) { return string.Equals(snapshot.Sqref, sqref, StringComparison.Ordinal); });
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

    private static void AssertConditionalFormatting(ConditionalFormattingSnapshot actual, string sqref, int ruleCount)
    {
        AssertEx.Equal(sqref, actual.Sqref);
        AssertEx.Equal(ruleCount, actual.Rules.Count);
    }

    private static string ResolveVisibility(string? state)
    {
        switch (state)
        {
            case "hidden":
                return "Hidden";
            case "veryHidden":
                return "VeryHidden";
            default:
                return "Visible";
        }
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

    private sealed record WorkbookSnapshot(
        bool UsesDate1904,
        bool HasSharedStrings,
        IReadOnlyList<SheetSnapshot> Sheets,
        IReadOnlyDictionary<string, CellSnapshot> Cells,
        IReadOnlyDictionary<string, StyleSnapshot> Styles,
        IReadOnlyDictionary<string, WorksheetSnapshot> Worksheets);

    private sealed record SheetSnapshot(string Name, string Visibility);
    private sealed record DefinedNameSnapshot(string Name, string Formula, int? LocalSheetIndex, bool Hidden, string Comment);
    private sealed record WorksheetSnapshot(
        string Name,
        string Visibility,
        string Dimension,
        IReadOnlyDictionary<int, RowSnapshot> Rows,
        IReadOnlyList<ColumnSnapshot> Columns,
        IReadOnlyList<string> MergeReferences,
        IReadOnlyList<HyperlinkSnapshot> Hyperlinks,
        IReadOnlyList<ValidationSnapshot> Validations,
        IReadOnlyList<ConditionalFormattingSnapshot> ConditionalFormattings,
        PageSetupSnapshot PageSetup);

    private sealed record RowSnapshot(double? Height, bool Hidden);
    private sealed record ColumnSnapshot(int MinColumnIndex, int MaxColumnIndex, double? Width, bool Hidden);
    private sealed record HyperlinkSnapshot(string Area, string Address, string Location, string ScreenTip, string Display);
    private sealed record ValidationSnapshot(string Sqref, string Type, string Operator, string ErrorStyle, bool AllowBlank, bool ShowDropDown, bool ShowInputMessage, bool ShowErrorMessage, string PromptTitle, string Prompt, string ErrorTitle, string Error, string Formula1, string Formula2);
    private sealed record ConditionalFormattingSnapshot(string Sqref, IReadOnlyList<ConditionalFormattingRuleSnapshot> Rules);
    private sealed record ConditionalFormattingRuleSnapshot(string Type, string Operator, string Priority, string StopIfTrue, string DxfId, IReadOnlyList<string> Formulas);
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










