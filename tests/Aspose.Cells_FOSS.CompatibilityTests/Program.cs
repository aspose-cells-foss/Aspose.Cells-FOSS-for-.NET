using Aspose.Cells_FOSS.Testing;

namespace Aspose.Cells_FOSS.CompatibilityTests;

internal static class Program
{
    private static int Main()
    {
        return TestRunner.Run(
            "CellData.CompatibilityTests",
            new TestCase("file_and_stream_load_paths_produce_same_values", FileAndStreamLoadPathsProduceSameValues),
            new TestCase("save_overloads_produce_equivalent_workbooks", SaveOverloadsProduceEquivalentWorkbooks),
            new TestCase("formula_setter_accepts_with_or_without_leading_equal", FormulaSetterAcceptsWithOrWithoutLeadingEqual),
            new TestCase("exception_mapping_uses_cells_exception_types", ExceptionMappingUsesCellsExceptionTypes),
            new TestCase("public_type_mapping_matches_after_roundtrip", PublicTypeMappingMatchesAfterRoundtrip),
            new TestCase("value_property_setter_matches_supported_scalar_behavior", ValuePropertySetterMatchesSupportedScalarBehavior),
            new TestCase("compatibility_members_follow_aspose_style", CompatibilityMembersFollowAsposeStyle),
            new TestCase("worksheet_view_members_follow_supported_patterns", WorksheetViewMembersFollowSupportedPatterns),
            new TestCase("worksheet_protection_members_follow_supported_patterns", WorksheetProtectionMembersFollowSupportedPatterns),
            new TestCase("autofilter_members_follow_supported_patterns", AutoFilterMembersFollowSupportedPatterns),
            new TestCase("file_and_stream_roundtrip_preserve_autofilter", FileAndStreamRoundtripPreserveAutoFilter),
            new TestCase("excel_input_autofilter_roundtrip_preserves_header_only_range", ExcelInputAutoFilterRoundtripPreservesHeaderOnlyRange),
            new TestCase("file_and_stream_roundtrip_preserve_defined_names", FileAndStreamRoundtripPreserveDefinedNames),
            new TestCase("file_and_stream_roundtrip_preserve_styles", FileAndStreamRoundtripPreserveStyles),
            new TestCase("file_and_stream_roundtrip_preserve_worksheet_settings", FileAndStreamRoundtripPreserveWorksheetSettings),
            new TestCase("file_and_stream_roundtrip_preserve_data_validations", FileAndStreamRoundtripPreserveDataValidations),
            new TestCase("file_and_stream_roundtrip_preserve_conditional_formattings", FileAndStreamRoundtripPreserveConditionalFormattings),
            new TestCase("file_and_stream_roundtrip_preserve_advanced_conditional_formattings", FileAndStreamRoundtripPreserveAdvancedConditionalFormattings),
            new TestCase("file_and_stream_roundtrip_preserve_page_setup", FileAndStreamRoundtripPreservePageSetup),
            new TestCase("workbook_metadata_members_follow_supported_patterns", WorkbookMetadataMembersFollowSupportedPatterns),
            new TestCase("file_and_stream_roundtrip_preserve_workbook_metadata", FileAndStreamRoundtripPreserveWorkbookMetadata));
    }

    private static void FileAndStreamLoadPathsProduceSameValues()
    {
        using var temp = new TemporaryDirectory("compat-load-paths");
        var path = temp.GetPath("book.xlsx");
        var workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
        workbook.Save(path);

        using var stream = File.OpenRead(path);
        var fromFile = new Workbook(path);
        var fromStream = new Workbook(stream);

        AssertWorkbookDataEquals(fromFile, fromStream);
    }

    private static void SaveOverloadsProduceEquivalentWorkbooks()
    {
        using var temp = new TemporaryDirectory("compat-save-overloads");
        var filePath = temp.GetPath("book-file.xlsx");
        var workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
        workbook.Save(filePath, SaveFormat.Xlsx);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);
        AssertWorkbookDataEquals(fromFile, fromStream);
    }

    private static void FormulaSetterAcceptsWithOrWithoutLeadingEqual()
    {
        var workbook = new Workbook();
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.PutValue(10);
        cell.Formula = "B1+C1";
        AssertEx.Equal("=B1+C1", cell.Formula);

        cell.Formula = "=D1+E1";
        AssertEx.Equal("=D1+E1", cell.Formula);
    }

    private static void ExceptionMappingUsesCellsExceptionTypes()
    {
        AssertEx.Throws<CellsException>(delegate { _ = new Workbook().Worksheets["missing"]; });
        AssertEx.Throws<CellsException>(delegate { _ = new Workbook().Worksheets[0].Cells["1A"]; });
        AssertEx.Throws<InvalidFileFormatException>(delegate { _ = new Workbook(new MemoryStream(new byte[] { 1, 2, 3, 4 })); });
    }

    private static void PublicTypeMappingMatchesAfterRoundtrip()
    {
        using var temp = new TemporaryDirectory("compat-types");
        var path = temp.GetPath("types.xlsx");
        var workbook = WorkbookScenarioFactory.CreateMixedCellWorkbook();
        workbook.Save(path);

        var loaded = new Workbook(path);
        AssertEx.Equal(CellValueType.String, loaded.Worksheets[0].Cells["A1"].Type);
        AssertEx.Equal(CellValueType.Number, loaded.Worksheets[0].Cells["B1"].Type);
        AssertEx.Equal(CellValueType.Boolean, loaded.Worksheets[0].Cells["C1"].Type);
        AssertEx.Equal(CellValueType.Number, loaded.Worksheets[0].Cells["D1"].Type);
        AssertEx.Equal(CellValueType.Number, loaded.Worksheets[0].Cells["E1"].Type);
        AssertEx.Equal(CellValueType.DateTime, loaded.Worksheets[0].Cells["F1"].Type);
        AssertEx.Equal(CellValueType.Formula, loaded.Worksheets[0].Cells["G1"].Type);
    }

    private static void ValuePropertySetterMatchesSupportedScalarBehavior()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "alpha";
        sheet.Cells["B1"].Value = 12;
        sheet.Cells["C1"].Value = true;
        sheet.Cells["D1"].Value = new DateTime(2024, 1, 2, 3, 4, 0, DateTimeKind.Unspecified);
        sheet.Cells["E1"].Value = null;

        AssertEx.Equal("alpha", sheet.Cells["A1"].Value as string);
        AssertEx.Equal(12, (int)sheet.Cells["B1"].Value!);
        AssertEx.Equal(true, (bool)sheet.Cells["C1"].Value!);
        AssertEx.Equal(CellValueType.DateTime, sheet.Cells["D1"].Type);
        AssertEx.Equal(string.Empty, sheet.Cells["E1"].DisplayStringValue);
    }

    private static void CompatibilityMembersFollowAsposeStyle()
    {
        using var workbook = new Workbook();
        workbook.Settings.Date1904 = true;
        AssertEx.True(workbook.Settings.Date1904);

        var index = workbook.Worksheets.Add();
        workbook.Worksheets[index].Name = "Report";
        workbook.Worksheets.ActiveSheetName = "Report";
        AssertEx.Equal(index, workbook.Worksheets.ActiveSheetIndex);

        var sheet = workbook.Worksheets[index];
        sheet.VisibilityType = VisibilityType.Hidden;
        AssertEx.Equal(VisibilityType.Hidden, sheet.VisibilityType);

        sheet.Cells["A1"].Value = "Docs";
        var hyperlinkIndex = sheet.Hyperlinks.Add(0, 0, 1, 1, "https://example.com/docs");
        var hyperlink = sheet.Hyperlinks[hyperlinkIndex];
        AssertEx.Equal(TargetModeType.External, hyperlink.LinkType);
        hyperlink.Delete();
        AssertEx.Equal(0, sheet.Hyperlinks.Count);

        var validationIndex = sheet.Validations.Add(CellArea.CreateCellArea("B2", "B2"));
        var validation = sheet.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.Formula1 = "\"Yes,No\"";
        AssertEx.Equal(ValidationType.List, sheet.Validations.GetValidationInCell(1, 1)!.Type);

        var conditionalFormattingIndex = sheet.ConditionalFormattings.Add();
        var conditionalFormatting = sheet.ConditionalFormattings[conditionalFormattingIndex];
        conditionalFormatting.AddArea(CellArea.CreateCellArea("C3", "C3"));
        conditionalFormatting.AddCondition(FormatConditionType.Expression, OperatorType.None, "C3>0", string.Empty);
        AssertEx.Equal(1, sheet.ConditionalFormattings.Count);

        var pageSetup = sheet.PageSetup;
        pageSetup.LeftMargin = 1.27d;
        AssertEx.Equal(0.5d, pageSetup.LeftMarginInch);
    }

    private static void WorksheetViewMembersFollowSupportedPatterns()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        sheet.TabColor = Color.FromArgb(255, 34, 68, 102);
        sheet.ShowGridlines = false;
        sheet.ShowRowColumnHeaders = false;
        sheet.ShowZeros = false;
        sheet.RightToLeft = true;
        sheet.Zoom = 85;

        AssertEx.Equal(Color.FromArgb(255, 34, 68, 102), sheet.TabColor);
        AssertEx.False(sheet.ShowGridlines);
        AssertEx.False(sheet.ShowRowColumnHeaders);
        AssertEx.False(sheet.ShowZeros);
        AssertEx.True(sheet.RightToLeft);
        AssertEx.Equal(85, sheet.Zoom);
    }

    private static void WorksheetProtectionMembersFollowSupportedPatterns()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        sheet.Protect();
        sheet.Protection.Objects = true;
        sheet.Protection.Scenarios = true;
        sheet.Protection.AutoFilter = true;
        sheet.Protection.SelectLockedCells = true;
        sheet.Protection.SelectUnlockedCells = true;

        AssertEx.True(sheet.Protection.IsProtected);
        AssertEx.True(sheet.Protection.Objects);
        AssertEx.True(sheet.Protection.Scenarios);
        AssertEx.True(sheet.Protection.AutoFilter);
        AssertEx.True(sheet.Protection.SelectLockedCells);
        AssertEx.True(sheet.Protection.SelectUnlockedCells);
    }

    private static void AutoFilterMembersFollowSupportedPatterns()
    {
        var workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
        AutoFilterScenarioFactory.AssertAutoFilter(workbook);
    }

    private static void FileAndStreamRoundtripPreserveAutoFilter()
    {
        using var temp = new TemporaryDirectory("compat-autofilter");
        var filePath = temp.GetPath("autofilter.xlsx");
        var workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        AutoFilterScenarioFactory.AssertAutoFilter(fromFile);
        AutoFilterScenarioFactory.AssertAutoFilter(fromStream);
    }

    private static void FileAndStreamRoundtripPreserveDefinedNames()
    {
        using var temp = new TemporaryDirectory("compat-defined-names");
        var filePath = temp.GetPath("defined-names.xlsx");
        var workbook = DefinedNameScenarioFactory.CreateDefinedNamesWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        DefinedNameScenarioFactory.AssertDefinedNames(fromFile);
        DefinedNameScenarioFactory.AssertDefinedNames(fromStream);
    }

    private static void FileAndStreamRoundtripPreserveStyles()
    {
        using var temp = new TemporaryDirectory("compat-styles");
        var filePath = temp.GetPath("styled-file.xlsx");
        var workbook = StyleScenarioFactory.CreateStyledWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        StyleScenarioFactory.AssertPrimaryStyle(fromFile.Worksheets[0].Cells["A1"].GetStyle());
        StyleScenarioFactory.AssertPrimaryStyle(fromStream.Worksheets[0].Cells["A1"].GetStyle());
        StyleScenarioFactory.AssertCustomNumberStyle(fromFile.Worksheets[0].Cells["B2"].GetStyle());
        StyleScenarioFactory.AssertCustomNumberStyle(fromStream.Worksheets[0].Cells["B2"].GetStyle());
        AssertEx.Equal(CellValueType.Blank, fromFile.Worksheets[0].Cells["B2"].Type);
        AssertEx.Equal(CellValueType.Blank, fromStream.Worksheets[0].Cells["B2"].Type);
    }

    private static void FileAndStreamRoundtripPreserveWorksheetSettings()
    {
        using var temp = new TemporaryDirectory("compat-worksheet-settings");
        var filePath = temp.GetPath("worksheet-settings.xlsx");
        var workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        WorksheetScenarioFactory.AssertWorksheetSettings(fromFile);
        WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(fromFile);
        WorksheetScenarioFactory.AssertWorksheetSettings(fromStream);
        WorksheetScenarioFactory.AssertWorksheetSettingsScenarioHasVisibleSheet(fromStream);
    }

    private static void FileAndStreamRoundtripPreserveDataValidations()
    {
        using var temp = new TemporaryDirectory("compat-validations");
        var filePath = temp.GetPath("validations.xlsx");
        var workbook = ValidationScenarioFactory.CreateValidationWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        ValidationScenarioFactory.AssertValidations(fromFile);
        ValidationScenarioFactory.AssertValidations(fromStream);
    }
    private static void FileAndStreamRoundtripPreserveConditionalFormattings()
    {
        using var temp = new TemporaryDirectory("compat-conditional-formatting");
        var filePath = temp.GetPath("conditional-formatting.xlsx");
        var workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        ConditionalFormattingScenarioFactory.AssertConditionalFormattings(fromFile);
        ConditionalFormattingScenarioFactory.AssertConditionalFormattings(fromStream);
    }
    private static void FileAndStreamRoundtripPreserveAdvancedConditionalFormattings()
    {
        using var temp = new TemporaryDirectory("compat-advanced-conditional-formatting");
        var filePath = temp.GetPath("advanced-conditional-formatting.xlsx");
        var workbook = ConditionalFormattingScenarioFactory.CreateAdvancedConditionalFormattingWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        ConditionalFormattingScenarioFactory.AssertAdvancedConditionalFormattings(fromFile);
        ConditionalFormattingScenarioFactory.AssertAdvancedConditionalFormattings(fromStream);
    }
    private static void AssertWorkbookDataEquals(Workbook expected, Workbook actual)
    {
        var expectedSheet = expected.Worksheets[0];
        var actualSheet = actual.Worksheets[0];

        foreach (var cellName in new[] { "A1", "B1", "C1", "D1", "E1", "F1", "G1" })
        {
            AssertEx.Equal(expectedSheet.Cells[cellName].Type, actualSheet.Cells[cellName].Type, $"Type mismatch for {cellName}.");
            AssertEx.Equal(expectedSheet.Cells[cellName].StringValue, actualSheet.Cells[cellName].StringValue, $"Value mismatch for {cellName}.");
            AssertEx.Equal(expectedSheet.Cells[cellName].Formula, actualSheet.Cells[cellName].Formula, $"Formula mismatch for {cellName}.");
        }
    }

    private static void FileAndStreamRoundtripPreservePageSetup()
    {
        using var temp = new TemporaryDirectory("compat-page-setup");
        var filePath = temp.GetPath("page-setup.xlsx");
        var workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        PageSetupScenarioFactory.AssertPageSetup(fromFile);
        PageSetupScenarioFactory.AssertPageSetup(fromStream);
    }

    private static void WorkbookMetadataMembersFollowSupportedPatterns()
    {
        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(workbook);
        AssertEx.Equal("Data", workbook.Worksheets[workbook.Properties.View.ActiveTab].Name);
    }

    private static void FileAndStreamRoundtripPreserveWorkbookMetadata()
    {
        using var temp = new TemporaryDirectory("compat-workbook-metadata");
        var filePath = temp.GetPath("workbook-metadata.xlsx");
        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        workbook.Save(filePath);

        using var stream = new MemoryStream();
        workbook.Save(stream, SaveFormat.Xlsx);
        stream.Position = 0;

        var fromFile = new Workbook(filePath);
        var fromStream = new Workbook(stream);

        WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(fromFile);
        WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(fromStream);
    
    }
    private static void ExcelInputAutoFilterRoundtripPreservesHeaderOnlyRange()
    {
        var inputPath = ResolveRepositoryFile("Input", "Autofilter.xlsx");
        var workbook = new Workbook(inputPath);
        AssertEx.Equal("A2:C2", workbook.Worksheets[0].AutoFilter.Range);
        AssertEx.Equal(0, workbook.Worksheets[0].AutoFilter.FilterColumns.Count);
        AssertEx.Equal(0, workbook.Worksheets[0].AutoFilter.SortState.SortConditions.Count);

        using var temp = new TemporaryDirectory("compat-input-autofilter");
        var roundtripPath = temp.GetPath("autofilter-roundtrip.xlsx");
        workbook.Save(roundtripPath);

        var reloaded = new Workbook(roundtripPath);
        AssertEx.Equal("A2:C2", reloaded.Worksheets[0].AutoFilter.Range);
        AssertEx.Equal(0, reloaded.Worksheets[0].AutoFilter.FilterColumns.Count);
        AssertEx.Equal(0, reloaded.Worksheets[0].AutoFilter.SortState.SortConditions.Count);

        var workbookXml = ZipPackageHelper.ReadEntryText(roundtripPath, "xl/workbook.xml");
        var worksheetXml = ZipPackageHelper.ReadEntryText(roundtripPath, "xl/worksheets/sheet1.xml");
        AssertEx.Contains(@"<autoFilter ref=""A2:C2""", worksheetXml);
        AssertEx.Contains(@"name=""_xlnm._FilterDatabase""", workbookXml);
        AssertEx.Contains(@"$A$2:$C$2", workbookXml);
    }

    private static string ResolveRepositoryFile(string firstSegment, string secondSegment)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Aspose.Cells_FOSS.sln")))
            {
                return Path.Combine(directory.FullName, firstSegment, secondSegment);
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Could not locate repository root.");
    }
}



