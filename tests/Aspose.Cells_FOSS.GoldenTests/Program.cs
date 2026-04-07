using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Testing;

namespace Aspose.Cells_FOSS.GoldenTests;

internal static class Program
{
    private static int Main()
    {
        return TestRunner.Run(
            "CellData.GoldenTests",
            new TestCase("xlsx_roundtrip_mixed_scalar_cells_file", XlsxRoundtripMixedScalarCellsFile),
            new TestCase("xlsx_roundtrip_mixed_scalar_cells_stream", XlsxRoundtripMixedScalarCellsStream),
            new TestCase("shared_strings_export_uses_sst_when_enabled", SharedStringsExportUsesSstWhenEnabled),
            new TestCase("inline_strings_export_when_shared_strings_disabled", InlineStringsExportWhenSharedStringsDisabled),
            new TestCase("formula_cells_persist_formula_and_cached_value", FormulaCellsPersistFormulaAndCachedValue),
            new TestCase("mac1904_datetime_roundtrip_and_markup", Mac1904DateTimeRoundtripAndMarkup),
            new TestCase("display_runtime_culture_after_roundtrip", DisplayRuntimeCultureAfterRoundtrip),
            new TestCase("styled_cells_roundtrip_and_emit_stylesheet", StyledCellsRoundtripAndEmitStylesheet),
            new TestCase("worksheet_settings_roundtrip_and_emit_expected_markup", WorksheetSettingsRoundtripAndEmitExpectedMarkup),
            new TestCase("worksheet_view_and_tab_color_roundtrip_and_emit_expected_markup", WorksheetViewAndTabColorRoundtripAndEmitExpectedMarkup),
            new TestCase("worksheet_protection_roundtrip_and_emit_expected_markup", WorksheetProtectionRoundtripAndEmitExpectedMarkup),
            new TestCase("autofilter_roundtrip_and_emit_expected_markup", AutoFilterRoundtripAndEmitExpectedMarkup),
            new TestCase("autofilter_omits_invalid_dxf_references", AutoFilterOmitsInvalidDxfReferences),
            new TestCase("defined_names_roundtrip_and_emit_expected_markup", DefinedNamesRoundtripAndEmitExpectedMarkup),
            new TestCase("worksheet_dimension_includes_merge_only_ranges", WorksheetDimensionIncludesMergeOnlyRanges),
            new TestCase("hyperlinks_roundtrip_and_emit_expected_markup", HyperlinksRoundtripAndEmitExpectedMarkup),
            new TestCase("data_validations_roundtrip_and_emit_expected_markup", DataValidationsRoundtripAndEmitExpectedMarkup),
            new TestCase("conditional_formattings_roundtrip_and_emit_expected_markup", ConditionalFormattingsRoundtripAndEmitExpectedMarkup),
            new TestCase("advanced_conditional_formattings_roundtrip_and_emit_expected_markup", AdvancedConditionalFormattingsRoundtripAndEmitExpectedMarkup),
            new TestCase("page_setup_roundtrip_and_emit_expected_markup", PageSetupRoundtripAndEmitExpectedMarkup),
            new TestCase("workbook_metadata_roundtrip_and_emit_expected_markup", WorkbookMetadataRoundtripAndEmitExpectedMarkup),
            new TestCase("extended_document_properties_do_not_inject_default_application_metadata", ExtendedDocumentPropertiesDoNotInjectDefaultApplicationMetadata),
            new TestCase("workbook_metadata_loads_from_root_relationship_targets", WorkbookMetadataLoadsFromRootRelationshipTargets),
            new TestCase("unreferenced_document_properties_parts_are_ignored", UnreferencedDocumentPropertiesPartsAreIgnored));
    }

    private static void XlsxRoundtripMixedScalarCellsFile()
    {
        using var temp = new TemporaryDirectory("golden-file");
        var path = temp.GetPath("mixed.xlsx");
        var workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
        workbook.Save(path);

        var loaded = new Workbook(path);
        AssertMixedWorkbook(loaded, false);
    }

    private static void XlsxRoundtripMixedScalarCellsStream()
    {
        var workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var loaded = new Workbook(stream);
        AssertMixedWorkbook(loaded, false);
    }

    private static void SharedStringsExportUsesSstWhenEnabled()
    {
        using var temp = new TemporaryDirectory("golden-sst");
        var path = temp.GetPath("shared-strings.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello");
        workbook.Worksheets[0].Cells["A2"].PutValue("Hello");
        workbook.Save(path, new SaveOptions { UseSharedStrings = true });

        AssertEx.True(ZipPackageHelper.EntryExists(path, "xl/sharedStrings.xml"));
        var sharedStringsXml = ZipPackageHelper.ReadEntryText(path, "xl/sharedStrings.xml");
        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");

        AssertEx.Contains("<sst", sharedStringsXml);
        AssertEx.Contains("uniqueCount=\"1\"", sharedStringsXml);
        AssertEx.Contains("t=\"s\"", worksheetXml);
    }

    private static void InlineStringsExportWhenSharedStringsDisabled()
    {
        using var temp = new TemporaryDirectory("golden-inline");
        var path = temp.GetPath("inline.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Inline");
        workbook.Save(path, new SaveOptions { UseSharedStrings = false });

        AssertEx.False(ZipPackageHelper.EntryExists(path, "xl/sharedStrings.xml"));
        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");

        AssertEx.Contains("t=\"inlineStr\"", worksheetXml);
        AssertEx.Contains("<t>Inline</t>", worksheetXml);
    }

    private static void FormulaCellsPersistFormulaAndCachedValue()
    {
        using var temp = new TemporaryDirectory("golden-formula");
        var path = temp.GetPath("formula.xlsx");
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].PutValue(30);
        sheet.Cells["C1"].Formula = "=A1+B1";
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains("<f>A1+B1</f>", worksheetXml);
        AssertEx.Contains("<v>30</v>", worksheetXml);

        var loaded = new Workbook(path);
        AssertEx.Equal("=A1+B1", loaded.Worksheets[0].Cells["C1"].Formula);
        AssertEx.Equal("30", loaded.Worksheets[0].Cells["C1"].StringValue);
    }

    private static void Mac1904DateTimeRoundtripAndMarkup()
    {
        using var temp = new TemporaryDirectory("golden-1904");
        var path = temp.GetPath("date1904.xlsx");
        var workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook(true);
        workbook.Save(path);

        var workbookXml = ZipPackageHelper.ReadEntryText(path, "xl/workbook.xml");
        AssertEx.Contains("date1904=\"1\"", workbookXml);

        var loaded = new Workbook(path);
        AssertEx.True(loaded.Settings.Date1904);
        var expected = new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc).Ticks;
        var actual = ((DateTime)loaded.Worksheets[0].Cells["F1"].Value!).Ticks;
        AssertEx.Equal(expected, actual);
    }

    private static void DisplayRuntimeCultureAfterRoundtrip()
    {
        using var temp = new TemporaryDirectory("golden-display-culture");
        var path = temp.GetPath("display-culture.xlsx");
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var timestamp = new DateTime(2024, 5, 6, 7, 8, 9);

        sheet.Cells["A1"].PutValue(1234.5d);
        var numericStyle = sheet.Cells["A1"].GetStyle();
        numericStyle.NumberFormat = "#,##0.00";
        sheet.Cells["A1"].SetStyle(numericStyle);

        sheet.Cells["B1"].PutValue(timestamp);
        var dateStyle = sheet.Cells["B1"].GetStyle();
        dateStyle.NumberFormat = "[$-409]dddd, mmmm d, yyyy";
        sheet.Cells["B1"].SetStyle(dateStyle);

        workbook.Save(path);

        var loaded = new Workbook(path);
        var frCulture = System.Globalization.CultureInfo.GetCultureInfo("fr-FR");
        loaded.Settings.Culture = frCulture;

        AssertEx.Equal(1234.5d.ToString("#,##0.00", frCulture), loaded.Worksheets[0].Cells["A1"].DisplayStringValue);
        AssertEx.Equal(timestamp.ToString("dddd, MMMM d, yyyy", System.Globalization.CultureInfo.GetCultureInfo("en-US")), loaded.Worksheets[0].Cells["B1"].DisplayStringValue);
    }

    private static void StyledCellsRoundtripAndEmitStylesheet()
    {
        using var temp = new TemporaryDirectory("golden-styles");
        var path = temp.GetPath("styles.xlsx");
        var workbook = StyleScenarioFactory.CreateStyledWorkbook();
        workbook.Save(path);

        AssertEx.True(ZipPackageHelper.EntryExists(path, "xl/styles.xml"));
        var stylesheetXml = ZipPackageHelper.ReadEntryText(path, "xl/styles.xml");
        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");

        AssertEx.Contains("Arial", stylesheetXml);
        AssertEx.Contains("0.0000", stylesheetXml);
        AssertEx.Contains("<strike", stylesheetXml);
        AssertEx.Contains("patternType=\"lightGrid\"", stylesheetXml);
        AssertEx.Contains("fgColor rgb=\"FFD2DC1E\"", stylesheetXml);
        AssertEx.Contains("bgColor rgb=\"FF0C2D4E\"", stylesheetXml);
        AssertEx.Contains("style=\"dotted\"", stylesheetXml);
        AssertEx.Contains("style=\"mediumDashDot\"", stylesheetXml);
        AssertEx.Contains("style=\"double\"", stylesheetXml);
        AssertEx.Contains("style=\"dashDotDot\"", stylesheetXml);
        AssertEx.Contains("style=\"slantDashDot\"", stylesheetXml);
        AssertEx.Contains("diagonalUp=\"1\"", stylesheetXml);
        AssertEx.Contains("diagonalDown=\"1\"", stylesheetXml);
        AssertEx.Contains("horizontal=\"distributed\"", stylesheetXml);
        AssertEx.Contains("vertical=\"distributed\"", stylesheetXml);
        AssertEx.Contains("indent=\"2\"", stylesheetXml);
        AssertEx.Contains("textRotation=\"45\"", stylesheetXml);
        AssertEx.Contains("shrinkToFit=\"1\"", stylesheetXml);
        AssertEx.Contains("readingOrder=\"2\"", stylesheetXml);
        AssertEx.Contains("relativeIndent=\"1\"", stylesheetXml);
        AssertEx.Contains("wrapText=\"1\"", stylesheetXml);
        AssertEx.Contains("locked=\"0\"", stylesheetXml);
        AssertEx.Contains("hidden=\"1\"", stylesheetXml);
        AssertEx.Contains("numFmtId=\"4\"", stylesheetXml);
        AssertFontElementOrder(stylesheetXml);
        AssertEx.Contains("s=\"", worksheetXml);

        var loaded = new Workbook(path);
        StyleScenarioFactory.AssertPrimaryStyle(loaded.Worksheets[0].Cells["A1"].GetStyle());
        StyleScenarioFactory.AssertCustomNumberStyle(loaded.Worksheets[0].Cells["B2"].GetStyle());
        AssertEx.Equal(CellValueType.Blank, loaded.Worksheets[0].Cells["B2"].Type);
        AssertEx.Equal(CellValueType.DateTime, loaded.Worksheets[0].Cells["C3"].Type);
    }

    private static void AssertFontElementOrder(string stylesheetXml)
    {
        var document = XDocument.Parse(stylesheetXml);
        XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var order = new Dictionary<string, int>(System.StringComparer.Ordinal)
        {
            { "b", 0 },
            { "i", 1 },
            { "strike", 2 },
            { "u", 3 },
            { "condense", 4 },
            { "extend", 5 },
            { "outline", 6 },
            { "shadow", 7 },
            { "charset", 8 },
            { "family", 9 },
            { "sz", 10 },
            { "color", 11 },
            { "vertAlign", 12 },
            { "scheme", 13 },
            { "name", 14 },
        };

        var fonts = document.Root?.Element(mainNs + "fonts")?.Elements(mainNs + "font") ?? Enumerable.Empty<XElement>();
        foreach (var font in fonts)
        {
            var previousOrder = -1;
            foreach (var child in font.Elements())
            {
                var childName = child.Name.LocalName;
                AssertEx.True(order.TryGetValue(childName, out var currentOrder), "Unexpected font child element: " + childName);
                AssertEx.True(currentOrder >= previousOrder, "Font child element order is invalid: " + font);
                previousOrder = currentOrder;
            }
        }
    }

    private static void WorksheetSettingsRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-worksheet-settings");
        var path = temp.GetPath("worksheet-settings.xlsx");
        var workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        var workbookXml = ZipPackageHelper.ReadEntryText(path, "xl/workbook.xml");
        AssertEx.Contains("dimension ref=\"A1:C4\"", worksheetXml);
        AssertEx.Contains("mergeCell ref=\"A1:B2\"", worksheetXml);
        AssertEx.Contains("hidden=\"1\"", worksheetXml);
        AssertEx.Contains("state=\"hidden\"", workbookXml);

        var loaded = new Workbook(path);
        WorksheetScenarioFactory.AssertWorksheetSettings(loaded);
        WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(loaded);
    }

    private static void WorksheetViewAndTabColorRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-worksheet-view");
        var path = temp.GetPath("worksheet-view.xlsx");
        var workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains("<sheetViews>", worksheetXml);
        AssertEx.Contains("showGridLines=\"0\"", worksheetXml);
        AssertEx.Contains("showRowColHeaders=\"0\"", worksheetXml);
        AssertEx.Contains("showZeros=\"0\"", worksheetXml);
        AssertEx.Contains("rightToLeft=\"1\"", worksheetXml);
        AssertEx.Contains("zoomScale=\"85\"", worksheetXml);
        AssertEx.Contains("tabColor rgb=\"FF224466\"", worksheetXml);

        var loaded = new Workbook(path);
        WorksheetScenarioFactory.AssertWorksheetSettings(loaded);
        WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(loaded);
    }

    private static void WorksheetProtectionRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-worksheet-protection");
        var path = temp.GetPath("worksheet-protection.xlsx");
        var workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains("<sheetProtection", worksheetXml);
        AssertEx.Contains("sheet=\"1\"", worksheetXml);
        AssertEx.Contains("objects=\"1\"", worksheetXml);
        AssertEx.Contains("scenarios=\"1\"", worksheetXml);
        AssertEx.Contains("formatCells=\"1\"", worksheetXml);
        AssertEx.Contains("insertRows=\"1\"", worksheetXml);
        AssertEx.Contains("autoFilter=\"1\"", worksheetXml);
        AssertEx.Contains("selectLockedCells=\"1\"", worksheetXml);
        AssertEx.Contains("selectUnlockedCells=\"1\"", worksheetXml);

        var loaded = new Workbook(path);
        WorksheetScenarioFactory.AssertWorksheetSettings(loaded);
        WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(loaded);
    }

    private static void AutoFilterRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-autofilter");
        var path = temp.GetPath("autofilter.xlsx");
        var workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        var workbookXml = ZipPackageHelper.ReadEntryText(path, "xl/workbook.xml");
        AssertEx.Contains(@"<autoFilter ref=""A1:E6""", worksheetXml);
        AssertEx.Contains(@"<filterColumn colId=""0"" hiddenButton=""1"">", worksheetXml);
        AssertEx.Contains(@"<filters><filter val=""Open"" /><filter val=""Closed"" /></filters>", worksheetXml);
        AssertEx.Contains(@"<customFilters and=""1"">", worksheetXml);
        AssertEx.Contains(@"operator=""greaterThanOrEqual""", worksheetXml);
        AssertEx.Contains(@"operator=""lessThanOrEqual""", worksheetXml);
        AssertEx.Contains(@"<colorFilter dxfId=""3"" cellColor=""1"" />", worksheetXml);
        AssertEx.Contains(@"<dynamicFilter type=""thisMonth"" val=""1"" maxVal=""31"" />", worksheetXml);
        AssertEx.Contains(@"<top10 top=""0"" percent=""1"" val=""10"" filterVal=""2.5"" />", worksheetXml);
        AssertEx.Contains(@"<sortState ref=""A2:E6"" caseSensitive=""1"" sortMethod=""pinYin"">", worksheetXml);
        AssertEx.Contains(@"<sortCondition ref=""B2:B6"" descending=""1"" sortBy=""value"" customList=""High,Medium,Low"" />", worksheetXml);
        AssertEx.Contains(@"<sortCondition ref=""C2:C6"" sortBy=""cellColor"" dxfId=""4"" />", worksheetXml);
        AssertEx.Contains(@"<sortCondition ref=""E2:E6"" sortBy=""icon"" iconSet=""3TrafficLights1"" iconId=""2"" />", worksheetXml);
        AssertEx.Contains(@"name=""_xlnm._FilterDatabase""", workbookXml);
        AssertEx.Contains(@"'Filtered'!$A$1:$E$6", workbookXml);

        var loaded = new Workbook(path);
        AutoFilterScenarioFactory.AssertAutoFilter(loaded);
    }


    private static void AutoFilterOmitsInvalidDxfReferences()
    {
        using var temp = new TemporaryDirectory("golden-autofilter-invalid-dxf");
        var path = temp.GetPath("autofilter-invalid-dxf.xlsx");
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "InvalidDxf";
        sheet.Cells["A1"].PutValue("Status");
        sheet.Cells["A2"].PutValue("Open");
        sheet.Cells["A3"].PutValue("Closed");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.AutoFilter.Range = "A1:B3";

        var column = sheet.AutoFilter.FilterColumns[sheet.AutoFilter.FilterColumns.Add(0)];
        column.ColorFilter.Enabled = true;
        column.ColorFilter.DifferentialStyleId = 9;
        column.ColorFilter.CellColor = true;

        sheet.AutoFilter.SortState.Ref = "A2:B3";
        var sortCondition = sheet.AutoFilter.SortState.SortConditions[sheet.AutoFilter.SortState.SortConditions.Add("B2:B3")];
        sortCondition.SortBy = "value";
        sortCondition.DifferentialStyleId = 8;

        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains(@"<autoFilter ref=""A1:B3""", worksheetXml);
        AssertEx.False(worksheetXml.Contains(@"<colorFilter", StringComparison.Ordinal));
        AssertEx.False(worksheetXml.Contains(@"dxfId=""8""", StringComparison.Ordinal));
        AssertEx.False(worksheetXml.Contains(@"dxfId=""9""", StringComparison.Ordinal));

        var loaded = new Workbook(path);
        AssertEx.Equal("A1:B3", loaded.Worksheets[0].AutoFilter.Range);
        AssertEx.Equal(0, loaded.Worksheets[0].AutoFilter.FilterColumns.Count);
        AssertEx.Equal(1, loaded.Worksheets[0].AutoFilter.SortState.SortConditions.Count);
        AssertEx.Null(loaded.Worksheets[0].AutoFilter.SortState.SortConditions[0].DifferentialStyleId);
    }
    private static void DefinedNamesRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-defined-names");
        var path = temp.GetPath("defined-names.xlsx");
        var workbook = DefinedNameScenarioFactory.CreateDefinedNamesWorkbook();
        workbook.Save(path);

        var workbookXml = ZipPackageHelper.ReadEntryText(path, "xl/workbook.xml");
        AssertEx.Contains("<definedNames>", workbookXml);
        AssertEx.Contains("name=\"GlobalRange\"", workbookXml);
        AssertEx.Contains("hidden=\"1\"", workbookXml);
        AssertEx.Contains("comment=\"Primary range\"", workbookXml);
        AssertEx.Contains("name=\"LocalCell\"", workbookXml);
        AssertEx.Contains("localSheetId=\"1\"", workbookXml);
        AssertEx.Contains("name=\"_xlnm.Print_Area\"", workbookXml);
        AssertEx.Contains("name=\"_xlnm.Print_Titles\"", workbookXml);

        var loaded = new Workbook(path);
        DefinedNameScenarioFactory.AssertDefinedNames(loaded);
    }

    private static void WorksheetDimensionIncludesMergeOnlyRanges()
    {
        using var temp = new TemporaryDirectory("golden-merge-dimension");
        var path = temp.GetPath("merge-only.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Name = "MergeOnly";
        workbook.Worksheets[0].Cells.Merge(4, 5, 2, 2);
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains("dimension ref=\"F5:G6\"", worksheetXml);
        AssertEx.Contains("mergeCell ref=\"F5:G6\"", worksheetXml);

        var loaded = new Workbook(path);
        AssertEx.Equal(1, loaded.Worksheets[0].Cells.MergedCells.Count);
        AssertEx.Equal(4, loaded.Worksheets[0].Cells.MergedCells[0].FirstRow);
        AssertEx.Equal(5, loaded.Worksheets[0].Cells.MergedCells[0].FirstColumn);
    }

    private static void HyperlinksRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-hyperlinks");
        var path = temp.GetPath("hyperlinks.xlsx");
        var workbook = HyperlinkScenarioFactory.CreateHyperlinkWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        var relsXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/_rels/sheet1.xml.rels");
        AssertEx.Contains("<hyperlinks>", worksheetXml);
        AssertEx.Contains("ref=\"A1\"", worksheetXml);
        AssertEx.Contains("location=\"'Target Sheet'!C3\"", worksheetXml);
        AssertEx.Contains("Target=\"https://example.com/docs?q=1\"", relsXml);
        AssertEx.Contains("Target=\"mailto:test@example.com\"", relsXml);

        var loaded = new Workbook(path);
        HyperlinkScenarioFactory.AssertHyperlinks(loaded);
    }

    private static void DataValidationsRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-validations");
        var path = temp.GetPath("validations.xlsx");
        var workbook = ValidationScenarioFactory.CreateValidationWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains("<dataValidations", worksheetXml);
        AssertEx.Contains("type=\"list\"", worksheetXml);
        AssertEx.Contains("type=\"decimal\"", worksheetXml);
        AssertEx.Contains("type=\"custom\"", worksheetXml);
        AssertEx.Contains("sqref=\"A1:A3\"", worksheetXml);
        AssertEx.Contains("sqref=\"B2:C3 E2:E3\"", worksheetXml);
        AssertEx.Contains("showDropDown=\"1\"", worksheetXml);
        AssertEx.Contains("<formula1>\"Open,Closed\"</formula1>", worksheetXml);
        AssertEx.Contains("<formula2>9.5</formula2>", worksheetXml);

        var loaded = new Workbook(path);
        ValidationScenarioFactory.AssertValidations(loaded);
    }
    private static void ConditionalFormattingsRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-conditional-formatting");
        var path = temp.GetPath("conditional-formatting.xlsx");
        var workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        var stylesXml = ZipPackageHelper.ReadEntryText(path, "xl/styles.xml");
        AssertEx.Contains("<conditionalFormatting", worksheetXml);
        AssertEx.Contains("type=\"cellIs\"", worksheetXml);
        AssertEx.Contains("type=\"expression\"", worksheetXml);
        AssertEx.Contains("operator=\"between\"", worksheetXml);
        AssertEx.Contains("stopIfTrue=\"1\"", worksheetXml);
        AssertEx.Contains("<formula>MOD(A1,2)=0</formula>", worksheetXml);
        AssertEx.Contains("<dxfs", stylesXml);
        AssertEx.Contains("count=\"3\"", stylesXml);

        var loaded = new Workbook(path);
        ConditionalFormattingScenarioFactory.AssertConditionalFormattings(loaded);
    }
    private static void AdvancedConditionalFormattingsRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-advanced-conditional-formatting");
        var path = temp.GetPath("advanced-conditional-formatting.xlsx");
        var workbook = ConditionalFormattingScenarioFactory.CreateAdvancedConditionalFormattingWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains("type=\"containsText\"", worksheetXml);
        AssertEx.Contains("type=\"notContainsText\"", worksheetXml);
        AssertEx.Contains("type=\"beginsWith\"", worksheetXml);
        AssertEx.Contains("type=\"endsWith\"", worksheetXml);
        AssertEx.Contains("type=\"timePeriod\"", worksheetXml);
        AssertEx.Contains("type=\"duplicateValues\"", worksheetXml);
        AssertEx.Contains("type=\"uniqueValues\"", worksheetXml);
        AssertEx.Contains("type=\"top10\"", worksheetXml);
        AssertEx.Contains("type=\"aboveAverage\"", worksheetXml);
        AssertEx.Contains("type=\"colorScale\"", worksheetXml);
        AssertEx.Contains("type=\"dataBar\"", worksheetXml);
        AssertEx.Contains("type=\"iconSet\"", worksheetXml);
        AssertEx.Contains("<colorScale>", worksheetXml);
        AssertEx.Contains("<dataBar>", worksheetXml);
        AssertEx.Contains("<iconSet iconSet=\"4Arrows\" reverse=\"1\" showValue=\"0\">", worksheetXml);

        var loaded = new Workbook(path);
        ConditionalFormattingScenarioFactory.AssertAdvancedConditionalFormattings(loaded);
    }
    private static void PageSetupRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-page-setup");
        var path = temp.GetPath("page-setup.xlsx");
        var workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
        workbook.Save(path);

        var worksheetXml = ZipPackageHelper.ReadEntryText(path, "xl/worksheets/sheet1.xml");
        AssertEx.Contains("pageMargins", worksheetXml);
        AssertEx.Contains("pageSetup", worksheetXml);
        AssertEx.Contains("orientation=\"landscape\"", worksheetXml);
        AssertEx.Contains("paperSize=\"9\"", worksheetXml);
        AssertEx.Contains("<rowBreaks", worksheetXml);
        AssertEx.Contains("<colBreaks", worksheetXml);

        var loaded = new Workbook(path);
        PageSetupScenarioFactory.AssertPageSetup(loaded);
    }

    private static void AssertMixedWorkbook(Workbook workbook, bool expectedDate1904)
    {
        AssertEx.Equal(expectedDate1904, workbook.Settings.Date1904);
        AssertEx.Equal("Data", workbook.Worksheets[0].Name);
        AssertEx.Equal("Hello", workbook.Worksheets[0].Cells["A1"].StringValue);
        AssertEx.Equal(123, (int)workbook.Worksheets[0].Cells["B1"].Value!);
        AssertEx.Equal(true, (bool)workbook.Worksheets[0].Cells["C1"].Value!);
        AssertEx.Equal(12.5m, (decimal)workbook.Worksheets[0].Cells["D1"].Value!);
        AssertEx.True(Math.Abs((double)workbook.Worksheets[0].Cells["E1"].Value! - 6.02214076E+23) < 1E+10);
        AssertEx.Equal(new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc).Ticks, ((DateTime)workbook.Worksheets[0].Cells["F1"].Value!).Ticks);
        AssertEx.Equal("=B1*2", workbook.Worksheets[0].Cells["G1"].Formula);
        AssertEx.Equal(20, (int)workbook.Worksheets[0].Cells["G1"].Value!);
    }

    private static void WorkbookMetadataRoundtripAndEmitExpectedMarkup()
    {
        using var temp = new TemporaryDirectory("golden-workbook-metadata");
        var path = temp.GetPath("workbook-metadata.xlsx");
        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        workbook.Save(path);

        var workbookXml = ZipPackageHelper.ReadEntryText(path, "xl/workbook.xml");
        var coreXml = ZipPackageHelper.ReadEntryText(path, "docProps/core.xml");
        var appXml = ZipPackageHelper.ReadEntryText(path, "docProps/app.xml");

        AssertEx.Contains("<workbookPr", workbookXml);
        AssertEx.Contains("codeName=\"WorkbookCode\"", workbookXml);
        AssertEx.Contains("showObjects=\"placeholders\"", workbookXml);
        AssertEx.Contains("<workbookProtection", workbookXml);
        AssertEx.Contains("workbookPassword=\"ABCD\"", workbookXml);
        AssertEx.Contains("<bookViews>", workbookXml);
        AssertEx.Contains("activeTab=\"1\"", workbookXml);
        AssertEx.Contains("showSheetTabs=\"0\"", workbookXml);
        AssertEx.Contains("<calcPr", workbookXml);
        AssertEx.Contains("calcMode=\"manual\"", workbookXml);
        AssertEx.Contains("refMode=\"R1C1\"", workbookXml);
        AssertEx.Contains("Quarterly Summary", coreXml);
        AssertEx.Contains("Automation", coreXml);
        AssertEx.Contains("Aspose.Cells_FOSS Tests", appXml);
        AssertEx.Contains("https://example.com/base/", appXml);

        var loaded = new Workbook(path);
        WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(loaded);
    }

    private static void ExtendedDocumentPropertiesDoNotInjectDefaultApplicationMetadata()
    {
        using var temp = new TemporaryDirectory("golden-extended-document-properties-no-default-app");
        var path = temp.GetPath("document-properties-no-default-app.xlsx");
        var workbook = new Workbook();
        workbook.DocumentProperties.Company = "Aspose Cells FOSS";
        workbook.DocumentProperties.Manager = "Release";
        workbook.Save(path);

        var appXml = ZipPackageHelper.ReadEntryText(path, "docProps/app.xml");
        AssertEx.Contains("<Company>Aspose Cells FOSS</Company>", appXml);
        AssertEx.Contains("<Manager>Release</Manager>", appXml);
        AssertEx.False(appXml.Contains("<Application>", StringComparison.Ordinal));
        AssertEx.False(appXml.Contains("<AppVersion>", StringComparison.Ordinal));

        var loaded = new Workbook(path);
        AssertEx.Equal("Aspose Cells FOSS", loaded.DocumentProperties.Company);
        AssertEx.Equal("Release", loaded.DocumentProperties.Manager);
        AssertEx.Equal(string.Empty, loaded.DocumentProperties.Extended.Application);
        AssertEx.Equal(string.Empty, loaded.DocumentProperties.Extended.AppVersion);
    }

    private static void WorkbookMetadataLoadsFromRootRelationshipTargets()
    {
        using var temp = new TemporaryDirectory("golden-workbook-metadata-root-relationships");
        var path = temp.GetPath("workbook-metadata-root-targets.xlsx");
        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        workbook.Save(path);

        ZipPackageHelper.MoveEntry(path, "docProps/core.xml", "metadata/core-props.xml");
        ZipPackageHelper.MoveEntry(path, "docProps/app.xml", "metadata/app-props.xml");
        RewriteDocumentPropertiesTargets(path, "/metadata/core-props.xml", "/metadata/app-props.xml");

        var loaded = new Workbook(path);
        WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(loaded);
    }

    private static void UnreferencedDocumentPropertiesPartsAreIgnored()
    {
        using var temp = new TemporaryDirectory("golden-workbook-metadata-orphaned-docprops");
        var path = temp.GetPath("workbook-metadata-orphaned-docprops.xlsx");
        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        workbook.Save(path);

        RemoveDocumentPropertiesRelationships(path);

        var loaded = new Workbook(path);
        AssertEx.Equal("WorkbookCode", loaded.Properties.CodeName);
        AssertDocumentPropertiesAreDefault(loaded);
    }

    private static void RewriteDocumentPropertiesTargets(string packagePath, string coreTarget, string appTarget)
    {
        ZipPackageHelper.RewriteXmlEntry(packagePath, "[Content_Types].xml", delegate(XDocument document)
        {
            var overrideName = XName.Get("Override", "http://schemas.openxmlformats.org/package/2006/content-types");
            foreach (var element in document.Root?.Elements(overrideName) ?? Enumerable.Empty<XElement>())
            {
                var partName = (string?)element.Attribute("PartName");
                if (string.Equals(partName, "/docProps/core.xml", StringComparison.Ordinal))
                {
                    element.SetAttributeValue("PartName", coreTarget);
                }
                else if (string.Equals(partName, "/docProps/app.xml", StringComparison.Ordinal))
                {
                    element.SetAttributeValue("PartName", appTarget);
                }
            }
        });

        ZipPackageHelper.RewriteXmlEntry(packagePath, "_rels/.rels", delegate(XDocument document)
        {
            var relationshipName = XName.Get("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships");
            foreach (var relationship in document.Root?.Elements(relationshipName) ?? Enumerable.Empty<XElement>())
            {
                var relationshipType = (string?)relationship.Attribute("Type");
                if (string.Equals(relationshipType, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", StringComparison.Ordinal))
                {
                    relationship.SetAttributeValue("Target", coreTarget.TrimStart('/'));
                }
                else if (string.Equals(relationshipType, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", StringComparison.Ordinal))
                {
                    relationship.SetAttributeValue("Target", appTarget.TrimStart('/'));
                }
            }
        });
    }

    private static void RemoveDocumentPropertiesRelationships(string packagePath)
    {
        ZipPackageHelper.RewriteXmlEntry(packagePath, "_rels/.rels", delegate(XDocument document)
        {
            var relationshipName = XName.Get("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships");
            var relationships = document.Root?.Elements(relationshipName).ToList() ?? new List<XElement>();
            for (var index = relationships.Count - 1; index >= 0; index--)
            {
                var relationship = relationships[index];
                var relationshipType = (string?)relationship.Attribute("Type");
                if (string.Equals(relationshipType, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", StringComparison.Ordinal)
                    || string.Equals(relationshipType, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", StringComparison.Ordinal))
                {
                    relationship.Remove();
                }
            }
        });
    }

    private static void AssertDocumentPropertiesAreDefault(Workbook workbook)
    {
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Title);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Subject);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Author);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Keywords);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Comments);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Category);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Company);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Manager);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Core.LastModifiedBy);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Core.Revision);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Core.ContentStatus);
        AssertEx.Null(workbook.DocumentProperties.Core.Created);
        AssertEx.Null(workbook.DocumentProperties.Core.Modified);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Extended.Application);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Extended.AppVersion);
        AssertEx.Equal(0, workbook.DocumentProperties.Extended.DocSecurity);
        AssertEx.Equal(string.Empty, workbook.DocumentProperties.Extended.HyperlinkBase);
        AssertEx.False(workbook.DocumentProperties.Extended.ScaleCrop);
        AssertEx.False(workbook.DocumentProperties.Extended.LinksUpToDate);
        AssertEx.False(workbook.DocumentProperties.Extended.SharedDoc);
    }
}




















