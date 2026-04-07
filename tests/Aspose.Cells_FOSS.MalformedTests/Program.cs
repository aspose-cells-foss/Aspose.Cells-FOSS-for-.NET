using System.Xml.Linq;
using Aspose.Cells_FOSS.Testing;

namespace Aspose.Cells_FOSS.MalformedTests;

internal static class Program
{
    private static int Main()
    {
        return TestRunner.Run(
            "CellData.MalformedTests",
            new TestCase("missing_workbook_xml_throws_invalid_file_format", MissingWorkbookXmlThrowsInvalidFileFormat),
            new TestCase("invalid_shared_string_index_records_lossy_recovery", InvalidSharedStringIndexRecordsLossyRecovery),
            new TestCase("invalid_cell_reference_is_skipped_with_warning", InvalidCellReferenceIsSkippedWithWarning),
            new TestCase("invalid_cell_reference_throws_in_strict_mode", InvalidCellReferenceThrowsInStrictMode),
            new TestCase("missing_sheetdata_synthesizes_empty_sheet", MissingSheetDataSynthesizesEmptySheet),
            new TestCase("missing_workbook_relationships_fall_back_by_convention", MissingWorkbookRelationshipsFallBackByConvention),
            new TestCase("invalid_date_serial_is_repaired_and_cleared", InvalidDateSerialIsRepairedAndCleared),
            new TestCase("invalid_style_index_records_warning_and_defaults_to_style_zero", InvalidStyleIndexRecordsWarningAndDefaultsToStyleZero),
            new TestCase("invalid_style_attributes_fall_back_to_safe_defaults", InvalidStyleAttributesFallBackToSafeDefaults),
            new TestCase("missing_dimension_records_recovery_and_preserves_settings", MissingDimensionRecordsRecoveryAndPreservesSettings),
            new TestCase("overlapping_merges_drop_conflicts_with_warning", OverlappingMergesDropConflictsWithWarning),
            new TestCase("invalid_column_span_throws_invalid_file_format", InvalidColumnSpanThrowsInvalidFileFormat),
            new TestCase("invalid_hyperlink_ref_is_dropped_with_warning", InvalidHyperlinkRefIsDroppedWithWarning),
            new TestCase("missing_hyperlink_relationship_records_warning", MissingHyperlinkRelationshipRecordsWarning),
            new TestCase("invalid_validation_sqref_is_dropped_with_warning", InvalidValidationSqrefIsDroppedWithWarning),
            new TestCase("invalid_validation_sqref_throws_in_strict_mode", InvalidValidationSqrefThrowsInStrictMode),
            new TestCase("overlapping_validations_keep_first", OverlappingValidationsKeepFirst),
            new TestCase("invalid_validation_type_falls_back_with_warning", InvalidValidationTypeFallsBackWithWarning),
            new TestCase("invalid_conditional_formatting_sqref_is_dropped_with_warning", InvalidConditionalFormattingSqrefIsDroppedWithWarning),
            new TestCase("unsupported_conditional_formatting_rule_is_dropped_with_warning", UnsupportedConditionalFormattingRuleIsDroppedWithWarning),
            new TestCase("invalid_conditional_formatting_dxfid_defaults_style", InvalidConditionalFormattingDxfIdDefaultsStyle),
            new TestCase("invalid_page_setup_attributes_record_warnings_and_fall_back", InvalidPageSetupAttributesRecordWarningsAndFallBack),
            new TestCase("invalid_sheet_view_attributes_record_warnings_and_fall_back", InvalidSheetViewAttributesRecordWarningsAndFallBack),
            new TestCase("invalid_sheet_protection_attributes_record_warnings_and_fall_back", InvalidSheetProtectionAttributesRecordWarningsAndFallBack),
            new TestCase("invalid_autofilter_ref_is_dropped_with_warning", InvalidAutoFilterRefIsDroppedWithWarning),
            new TestCase("invalid_autofilter_sortstate_records_warning", InvalidAutoFilterSortStateRecordsWarning),
            new TestCase("invalid_autofilter_dxfids_record_warning_and_drop_invalid_metadata", InvalidAutoFilterDxfIdsRecordWarningAndDropInvalidMetadata),
            new TestCase("invalid_defined_name_scope_records_warning", InvalidDefinedNameScopeRecordsWarning),
            new TestCase("invalid_print_titles_defined_name_records_warning", InvalidPrintTitlesDefinedNameRecordsWarning),
            new TestCase("invalid_workbook_metadata_attributes_record_warning", InvalidWorkbookMetadataAttributesRecordWarning),
            new TestCase("invalid_document_properties_record_warning", InvalidDocumentPropertiesRecordWarning));
    }

    private static void MissingWorkbookXmlThrowsInvalidFileFormat()
    {
        using var temp = new TemporaryDirectory("malformed-missing-workbook");
        var path = temp.GetPath("broken.xlsx");
        ZipPackageHelper.CreatePackage(path, new Dictionary<string, string>());

        AssertEx.Throws<InvalidFileFormatException>(delegate { _ = new Workbook(path); });
    }

    private static void InvalidSharedStringIndexRecordsLossyRecovery()
    {
        using var temp = new TemporaryDirectory("malformed-sst");
        var path = temp.GetPath("sst.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Hello");
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("<v>0</v>", "<v>99</v>"); });

        var loaded = new Workbook(path);
        AssertEx.Equal(string.Empty, loaded.Worksheets[0].Cells["A1"].StringValue);
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "SST-L001"; }));
    }

    private static void InvalidCellReferenceIsSkippedWithWarning()
    {
        using var temp = new TemporaryDirectory("malformed-ref-warning");
        var path = temp.GetPath("invalid-ref.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(42);
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("r=\"A1\"", "r=\"1A\""); });

        var loaded = new Workbook(path);
        AssertEx.Equal(CellValueType.Blank, loaded.Worksheets[0].Cells["A1"].Type);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "CELL-F001"; }));
    }

    private static void InvalidCellReferenceThrowsInStrictMode()
    {
        using var temp = new TemporaryDirectory("malformed-ref-strict");
        var path = temp.GetPath("invalid-ref-strict.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(42);
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("r=\"A1\"", "r=\"1A\""); });

        AssertEx.Throws<InvalidFileFormatException>(delegate { _ = new Workbook(path, new LoadOptions { StrictMode = true }); });
    }

    private static void MissingSheetDataSynthesizesEmptySheet()
    {
        using var temp = new TemporaryDirectory("malformed-sheetdata");
        var path = temp.GetPath("missing-sheetdata.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(42);
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            document.Root?.Element(XName.Get("sheetData", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?.Remove();
        });

        var loaded = new Workbook(path);
        AssertEx.Equal(CellValueType.Blank, loaded.Worksheets[0].Cells["A1"].Type);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "ACF-WS-001"; }));
    }

    private static void MissingWorkbookRelationshipsFallBackByConvention()
    {
        using var temp = new TemporaryDirectory("malformed-rels");
        var path = temp.GetPath("missing-rels.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(42);
        workbook.Save(path);

        ZipPackageHelper.DeleteEntry(path, "xl/_rels/workbook.xml.rels");

        var loaded = new Workbook(path);
        AssertEx.Equal("42", loaded.Worksheets[0].Cells["A1"].StringValue);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "PKG-R001"; }));
    }

    private static void InvalidDateSerialIsRepairedAndCleared()
    {
        using var temp = new TemporaryDirectory("malformed-date");
        var path = temp.GetPath("invalid-date.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue(new DateTime(2024, 5, 6, 7, 8, 9));
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml)
        {
            var start = xml.IndexOf("<v>", StringComparison.Ordinal);
            var end = xml.IndexOf("</v>", StringComparison.Ordinal);
            return start >= 0 && end > start ? xml.Substring(0, start + 3) + "oops" + xml.Substring(end) : xml;
        });

        var loaded = new Workbook(path);
        AssertEx.Equal(CellValueType.Blank, loaded.Worksheets[0].Cells["A1"].Type);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "CELL-R002"; }));
    }

    private static void InvalidStyleIndexRecordsWarningAndDefaultsToStyleZero()
    {
        using var temp = new TemporaryDirectory("malformed-style-index");
        var path = temp.GetPath("invalid-style.xlsx");
        var workbook = StyleScenarioFactory.CreateStyledWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("s=\"1\"", "s=\"999\""); });

        var loaded = new Workbook(path);
        AssertEx.Equal("1234.567", loaded.Worksheets[0].Cells["A1"].StringValue);
        var style = loaded.Worksheets[0].Cells["A1"].GetStyle();
        AssertEx.False(style.Font.Bold);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "STYLE-F001"; }));
    }

    private static void InvalidStyleAttributesFallBackToSafeDefaults()
    {
        using var temp = new TemporaryDirectory("malformed-style-attributes");
        var path = temp.GetPath("invalid-style-attributes.xlsx");
        var workbook = StyleScenarioFactory.CreateStyledWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/styles.xml", delegate(XDocument document)
        {
            var ns = XName.Get("cellXfs", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var xf = document.Root?
                .Element(ns)?
                .Elements(XName.Get("xf", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                .Skip(1)
                .FirstOrDefault();
            var alignment = xf?.Element(XName.Get("alignment", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
            alignment?.SetAttributeValue("indent", "-5");
            alignment?.SetAttributeValue("textRotation", "999");
            alignment?.SetAttributeValue("readingOrder", "42");

            var border = document.Root?
                .Element(XName.Get("borders", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?
                .Elements(XName.Get("border", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                .Skip(1)
                .FirstOrDefault();
            border?.Element(XName.Get("left", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?.SetAttributeValue("style", "mystery");
        });

        var loaded = new Workbook(path);
        var style = loaded.Worksheets[0].Cells["A1"].GetStyle();
        AssertEx.Equal(0, style.IndentLevel);
        AssertEx.Equal(0, style.TextRotation);
        AssertEx.Equal(0, style.ReadingOrder);
        AssertEx.Equal(BorderStyleType.None, style.Borders.Left.LineStyle);
    }
    private static void MissingDimensionRecordsRecoveryAndPreservesSettings()
    {
        using var temp = new TemporaryDirectory("malformed-missing-dimension");
        var path = temp.GetPath("missing-dimension.xlsx");
        var workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            document.Root?.Element(XName.Get("dimension", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?.Remove();
        });

        var loaded = new Workbook(path);
        WorksheetScenarioFactory.AssertWorksheetSettings(loaded);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WS-R001"; }));
    }

    private static void OverlappingMergesDropConflictsWithWarning()
    {
        using var temp = new TemporaryDirectory("malformed-overlapping-merges");
        var path = temp.GetPath("overlap-merges.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells.Merge(0, 0, 2, 2);
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("</mergeCells>", "<mergeCell ref=\"B2:C3\"/></mergeCells>"); });

        var loaded = new Workbook(path);
        AssertEx.Equal(1, loaded.Worksheets[0].Cells.MergedCells.Count);
        AssertEx.Equal(0, loaded.Worksheets[0].Cells.MergedCells[0].FirstRow);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "MRG-L001"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void InvalidColumnSpanThrowsInvalidFileFormat()
    {
        using var temp = new TemporaryDirectory("malformed-column-span");
        var path = temp.GetPath("invalid-column-span.xlsx");
        var workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("min=\"1\" max=\"1\"", "min=\"5\" max=\"2\""); });

        AssertEx.Throws<InvalidFileFormatException>(delegate { _ = new Workbook(path); });
    }

    private static void InvalidHyperlinkRefIsDroppedWithWarning()
    {
        using var temp = new TemporaryDirectory("malformed-hyperlink-ref");
        var path = temp.GetPath("invalid-hyperlink-ref.xlsx");
        var workbook = HyperlinkScenarioFactory.CreateHyperlinkWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            var hyperlink = document.Root?
                .Element(XName.Get("hyperlinks", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?
                .Elements(XName.Get("hyperlink", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                .FirstOrDefault();
            hyperlink?.SetAttributeValue("ref", "1A");
        });

        var loaded = new Workbook(path);
        AssertEx.Equal(2, loaded.Worksheets[0].Hyperlinks.Count);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "HL-L001"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void MissingHyperlinkRelationshipRecordsWarning()
    {
        using var temp = new TemporaryDirectory("malformed-hyperlink-rels");
        var path = temp.GetPath("missing-hyperlink-rels.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Docs");
        workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://example.com/docs");
        workbook.Save(path);

        ZipPackageHelper.DeleteEntry(path, "xl/worksheets/_rels/sheet1.xml.rels");

        var loaded = new Workbook(path);
        AssertEx.Equal(0, loaded.Worksheets[0].Hyperlinks.Count);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "HL-L002"; }));
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "HL-L004"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void InvalidValidationSqrefIsDroppedWithWarning()
    {
        using var temp = new TemporaryDirectory("malformed-validation-sqref");
        var path = temp.GetPath("invalid-validation-sqref.xlsx");
        var workbook = ValidationScenarioFactory.CreateValidationWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            var validation = document.Root?
                .Element(XName.Get("dataValidations", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?
                .Elements(XName.Get("dataValidation", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                .FirstOrDefault();
            validation?.SetAttributeValue("sqref", "1A");
        });

        var loaded = new Workbook(path);
        AssertEx.Equal(2, loaded.Worksheets[0].Validations.Count);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "DV-L001"; }));
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "DV-L002"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void InvalidValidationSqrefThrowsInStrictMode()
    {
        using var temp = new TemporaryDirectory("malformed-validation-sqref-strict");
        var path = temp.GetPath("invalid-validation-sqref-strict.xlsx");
        var workbook = ValidationScenarioFactory.CreateValidationWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("sqref=\"A1:A3\"", "sqref=\"1A\""); });

        AssertEx.Throws<InvalidFileFormatException>(delegate { _ = new Workbook(path, new LoadOptions { StrictMode = true }); });
    }

    private static void OverlappingValidationsKeepFirst()
    {
        using var temp = new TemporaryDirectory("malformed-validation-overlap");
        var path = temp.GetPath("overlap-validations.xlsx");
        var workbook = ValidationScenarioFactory.CreateValidationWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml)
        {
            var marker = "</dataValidations>";
            var injected = "<dataValidation type=\"whole\" sqref=\"A2:A4\"><formula1>1</formula1><formula2>2</formula2></dataValidation>";
            return xml.Replace(marker, injected + marker);
        });

        var loaded = new Workbook(path);
        AssertEx.Equal(3, loaded.Worksheets[0].Validations.Count);
        AssertEx.Equal(ValidationType.List, loaded.Worksheets[0].Validations.GetValidationInCell(1, 0)!.Type);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "DV-L003"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void InvalidValidationTypeFallsBackWithWarning()
    {
        using var temp = new TemporaryDirectory("malformed-validation-type");
        var path = temp.GetPath("invalid-validation-type.xlsx");
        var workbook = ValidationScenarioFactory.CreateValidationWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("type=\"list\"", "type=\"mystery\""); });

        var loaded = new Workbook(path);
        AssertEx.Equal(ValidationType.AnyValue, loaded.Worksheets[0].Validations[0].Type);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "DV-R001"; }));
    }
    private static void InvalidConditionalFormattingSqrefIsDroppedWithWarning()
    {
        using var temp = new TemporaryDirectory("malformed-conditional-formatting-sqref");
        var path = temp.GetPath("invalid-conditional-formatting-sqref.xlsx");
        var workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("sqref=\"A1:A5\"", "sqref=\"1A\""); });

        var loaded = new Workbook(path);
        AssertEx.Equal(1, loaded.Worksheets[0].ConditionalFormattings.Count);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "CF-L001"; }));
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "CF-L003"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void UnsupportedConditionalFormattingRuleIsDroppedWithWarning()
    {
        using var temp = new TemporaryDirectory("malformed-conditional-formatting-type");
        var path = temp.GetPath("unsupported-conditional-formatting-type.xlsx");
        var workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("type=\"expression\"", "type=\"containsBlanks\""); });

        var loaded = new Workbook(path);
        AssertEx.Equal(2, loaded.Worksheets[0].ConditionalFormattings.Count);
        AssertEx.Equal(1, loaded.Worksheets[0].ConditionalFormattings[0].Count);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "CF-L002"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void InvalidConditionalFormattingDxfIdDefaultsStyle()
    {
        using var temp = new TemporaryDirectory("malformed-conditional-formatting-dxfid");
        var path = temp.GetPath("invalid-conditional-formatting-dxfid.xlsx");
        var workbook = ConditionalFormattingScenarioFactory.CreateConditionalFormattingWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteEntryText(path, "xl/worksheets/sheet1.xml", delegate(string xml) { return xml.Replace("dxfId=\"0\"", "dxfId=\"999\""); });

        var loaded = new Workbook(path);
        var collection = loaded.Worksheets[0].ConditionalFormattings[0];
        AssertEx.Equal(FillPattern.None, collection[0].Style.Pattern);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "CF-R002"; }));
    }
    private static void InvalidPageSetupAttributesRecordWarningsAndFallBack()
    {
        using var temp = new TemporaryDirectory("malformed-page-setup");
        var path = temp.GetPath("page-setup-invalid.xlsx");
        var workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            var ns = XName.Get("pageSetup", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var pageSetup = document.Root?.Element(ns);
            pageSetup?.SetAttributeValue("scale", "999");
            pageSetup?.SetAttributeValue("orientation", "sideways");
            var pageMargins = document.Root?.Element(XName.Get("pageMargins", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
            pageMargins?.SetAttributeValue("left", "-1");
        });

        var loaded = new Workbook(path);
        var pageSetup = loaded.Worksheets[0].PageSetup;
        AssertEx.Equal(95, pageSetup.Scale ?? 95);
        AssertEx.Equal(PageOrientationType.Default, pageSetup.Orientation);
        AssertEx.Equal(0.7d, pageSetup.LeftMarginInch);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "PG-L001"; }));
        AssertEx.True(loaded.LoadDiagnostics.Issues.Count(delegate(LoadIssue issue) { return issue.Code == "PG-L002"; }) >= 1);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "PG-L003"; }));
    }

    private static void InvalidSheetProtectionAttributesRecordWarningsAndFallBack()
    {
        using var temp = new TemporaryDirectory("malformed-sheet-protection");
        var path = temp.GetPath("invalid-sheet-protection.xlsx");
        var workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            var protection = document.Root?.Element(XName.Get("sheetProtection", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
            protection?.SetAttributeValue("sheet", "maybe");
            protection?.SetAttributeValue("objects", "oops");
            protection?.SetAttributeValue("formatCells", "1");
            protection?.SetAttributeValue("selectUnlockedCells", "1");
        });

        var loaded = new Workbook(path);
        var sheet = loaded.Worksheets[0];
        AssertEx.True(sheet.Protection.IsProtected);
        AssertEx.False(sheet.Protection.Objects);
        AssertEx.True(sheet.Protection.FormatCells);
        AssertEx.True(sheet.Protection.SelectUnlockedCells);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Count(delegate(LoadIssue issue) { return issue.Code == "WS-L009"; }) >= 2);
    }

    private static void InvalidAutoFilterRefIsDroppedWithWarning()
    {
        using var temp = new TemporaryDirectory("malformed-autofilter-ref");
        var path = temp.GetPath("invalid-autofilter-ref.xlsx");
        var workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            document.Root?.Element(XName.Get("autoFilter", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?.SetAttributeValue("ref", "1A");
        });

        var loaded = new Workbook(path);
        AssertEx.Equal(string.Empty, loaded.Worksheets[0].AutoFilter.Range);
        AssertEx.Equal(0, loaded.Worksheets[0].AutoFilter.FilterColumns.Count);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WS-L010"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void InvalidAutoFilterSortStateRecordsWarning()
    {
        using var temp = new TemporaryDirectory("malformed-autofilter-sortstate");
        var path = temp.GetPath("invalid-autofilter-sortstate.xlsx");
        var workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            var autoFilter = document.Root?.Element(XName.Get("autoFilter", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
            autoFilter?.Element(XName.Get("sortState", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?.SetAttributeValue("ref", "1A");
            autoFilter?.Add(new XElement(XName.Get("filterColumn", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
                new XAttribute("colId", "bad")));
        });

        var loaded = new Workbook(path);
        AssertEx.Equal("A1:E6", loaded.Worksheets[0].AutoFilter.Range);
        AssertEx.Equal(5, loaded.Worksheets[0].AutoFilter.FilterColumns.Count);
        AssertEx.Equal(string.Empty, loaded.Worksheets[0].AutoFilter.SortState.Ref);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WS-L011"; }));
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WS-L012"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }


    private static void InvalidAutoFilterDxfIdsRecordWarningAndDropInvalidMetadata()
    {
        using var temp = new TemporaryDirectory("malformed-autofilter-dxfids");
        var path = temp.GetPath("invalid-autofilter-dxfids.xlsx");
        var workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            var autoFilterName = XName.Get("autoFilter", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var filterColumnName = XName.Get("filterColumn", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var colorFilterName = XName.Get("colorFilter", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var sortStateName = XName.Get("sortState", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var sortConditionName = XName.Get("sortCondition", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var autoFilter = document.Root?.Element(autoFilterName);
            autoFilter?.Elements(filterColumnName)
                .FirstOrDefault(delegate(XElement element) { return (string?)element.Attribute("colId") == "2"; })
                ?.Element(colorFilterName)
                ?.SetAttributeValue("dxfId", "99");
            autoFilter?.Element(sortStateName)
                ?.Element(sortConditionName)
                ?.SetAttributeValue("dxfId", "99");
        });

        var loaded = new Workbook(path);
        var colorColumn = loaded.Worksheets[0].AutoFilter.FilterColumns[2];
        AssertEx.False(colorColumn.ColorFilter.Enabled);
        AssertEx.Null(colorColumn.ColorFilter.DifferentialStyleId);
        AssertEx.True(loaded.Worksheets[0].AutoFilter.SortState.SortConditions.Count > 0);
        AssertEx.Null(loaded.Worksheets[0].AutoFilter.SortState.SortConditions[0].DifferentialStyleId);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WS-L011"; }));
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WS-L012"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }
    private static void InvalidDefinedNameScopeRecordsWarning()
    {
        using var temp = new TemporaryDirectory("malformed-defined-name-scope");
        var path = temp.GetPath("invalid-defined-name-scope.xlsx");
        var workbook = DefinedNameScenarioFactory.CreateDefinedNamesWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/workbook.xml", delegate(XDocument document)
        {
            var ns = XName.Get("definedNames", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var definedNames = document.Root?.Element(ns);
            var localName = definedNames?.Elements(XName.Get("definedName", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                .FirstOrDefault(delegate(XElement element) { return string.Equals((string?)element.Attribute("name"), "LocalCell", StringComparison.Ordinal); });
            if (localName is not null)
            {
                localName.SetAttributeValue("localSheetId", "99");
            }
        });

        var loaded = new Workbook(path);
        AssertEx.Equal(1, loaded.DefinedNames.Count);
        AssertEx.Equal("GlobalRange", loaded.DefinedNames[0].Name);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WB-L002"; }));
        AssertEx.True(loaded.LoadDiagnostics.HasDataLossRisk);
    }

    private static void InvalidPrintTitlesDefinedNameRecordsWarning()
    {
        using var temp = new TemporaryDirectory("malformed-print-titles");
        var path = temp.GetPath("invalid-print-titles.xlsx");
        var workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/workbook.xml", delegate(XDocument document)
        {
            var ns = XName.Get("definedNames", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var definedNames = document.Root?.Element(ns);
            var printTitles = definedNames?.Elements(XName.Get("definedName", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                .FirstOrDefault(delegate(XElement element) { return string.Equals((string?)element.Attribute("name"), "_xlnm.Print_Titles", StringComparison.Ordinal); });
            if (printTitles is not null)
            {
                printTitles.Value = "'Print Sheet'!$A$1";
            }
        });

        var loaded = new Workbook(path);
        AssertEx.Equal(string.Empty, loaded.Worksheets[0].PageSetup.PrintTitleRows);
        AssertEx.Equal(string.Empty, loaded.Worksheets[0].PageSetup.PrintTitleColumns);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "PG-L004"; }));
    }

    private static void InvalidSheetViewAttributesRecordWarningsAndFallBack()
    {
        using var temp = new TemporaryDirectory("malformed-sheet-view");
        var path = temp.GetPath("invalid-sheet-view.xlsx");
        var workbook = WorksheetScenarioFactory.CreateWorksheetSettingsWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/worksheets/sheet1.xml", delegate(XDocument document)
        {
            var ns = XName.Get("sheetViews", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var sheetView = document.Root?.Element(ns)?.Element(XName.Get("sheetView", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
            sheetView?.SetAttributeValue("zoomScale", "999");
            sheetView?.SetAttributeValue("showGridLines", "maybe");
            var tabColor = document.Root?.Element(XName.Get("sheetPr", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?.Element(XName.Get("tabColor", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
            tabColor?.SetAttributeValue("rgb", "broken");
        });

        var loaded = new Workbook(path);
        var sheet = loaded.Worksheets[0];
        AssertEx.Equal(Color.Empty, sheet.TabColor);
        AssertEx.True(sheet.ShowGridlines);
        AssertEx.False(sheet.ShowRowColumnHeaders);
        AssertEx.False(sheet.ShowZeros);
        AssertEx.True(sheet.RightToLeft);
        AssertEx.Equal(100, sheet.Zoom);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WS-L007"; }));
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WS-L008"; }));
    }

    private static void InvalidWorkbookMetadataAttributesRecordWarning()
    {
        using var temp = new TemporaryDirectory("malformed-workbook-metadata");
        var path = temp.GetPath("workbook-metadata.xlsx");
        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "xl/workbook.xml", delegate(XDocument document)
        {
            var ns = XName.Get("workbookPr", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            document.Root?.Element(ns)?.SetAttributeValue("showObjects", "mystery");
            document.Root?.Element(XName.Get("bookViews", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?
                .Element(XName.Get("workbookView", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?
                .SetAttributeValue("activeTab", "99");
            document.Root?.Element(XName.Get("calcPr", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))?
                .SetAttributeValue("iterateDelta", "oops");
        });

        var loaded = new Workbook(path);
        AssertEx.Equal("all", loaded.Properties.ShowObjects);
        AssertEx.Equal(0.001d, loaded.Properties.Calculation.IterateDelta);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WB-L003"; }));
    }

    private static void InvalidDocumentPropertiesRecordWarning()
    {
        using var temp = new TemporaryDirectory("malformed-document-properties");
        var path = temp.GetPath("document-properties.xlsx");
        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        workbook.Save(path);

        ZipPackageHelper.RewriteXmlEntry(path, "docProps/core.xml", delegate(XDocument document)
        {
            document.Root?.Element(XName.Get("created", "http://purl.org/dc/terms/"))?.SetValue("not-a-date");
        });

        var loaded = new Workbook(path);
        AssertEx.Null(loaded.DocumentProperties.Core.Created);
        AssertEx.True(loaded.LoadDiagnostics.Issues.Any(delegate(LoadIssue issue) { return issue.Code == "WB-L004"; }));
    }}












