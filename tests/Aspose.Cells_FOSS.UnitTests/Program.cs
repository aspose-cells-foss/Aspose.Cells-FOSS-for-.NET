using System.Globalization;
using Aspose.Cells_FOSS.Core;
using Aspose.Cells_FOSS.Testing;

namespace Aspose.Cells_FOSS.UnitTests;

internal static class Program
{
    private static int Main()
    {
        return TestRunner.Run(
            "CellData.UnitTests",
            new TestCase("a1_indexers_roundtrip", A1IndexersRoundtrip),
            new TestCase("put_value_overloads_assign_expected_types", PutValueOverloadsAssignExpectedTypes),
            new TestCase("stringvalue_formats_supported_scalar_types", StringValueFormatsSupportedScalarTypes),
            new TestCase("displaystringvalue_applies_numeric_formats_and_stringvalue_stays_raw", DisplayStringValueAppliesNumericFormatsAndStringValueStaysRaw),
            new TestCase("displaystringvalue_applies_date_and_text_formats", DisplayStringValueAppliesDateAndTextFormats),
            new TestCase("displaystringvalue_uses_workbook_culture_and_locale_directives", DisplayStringValueUsesWorkbookCultureAndLocaleDirectives),
            new TestCase("displaystringvalue_applies_extended_date_tokens", DisplayStringValueAppliesExtendedDateTokens),
            new TestCase("displaystringvalue_applies_long_time_and_elapsed_fraction_with_culture", DisplayStringValueAppliesLongTimeAndElapsedFractionWithCulture),
            new TestCase("formula_property_normalizes_and_preserves_cached_value", FormulaPropertyNormalizesAndPreservesCachedValue),
            new TestCase("blank_cells_are_blank_by_default", BlankCellsAreBlankByDefault),
            new TestCase("worksheet_name_and_collection_guards", WorksheetNameAndCollectionGuards),
            new TestCase("style_mutation_requires_setstyle_and_returns_clones", StyleMutationRequiresSetStyleAndReturnsClones),
            new TestCase("style_api_covers_all_public_settings", StyleApiCoversAllPublicSettings),
            new TestCase("worksheet_row_column_and_merge_apis_mutate_expected_settings", WorksheetRowColumnAndMergeApisMutateExpectedSettings),
            new TestCase("worksheet_view_apis_mutate_expected_settings", WorksheetViewApisMutateExpectedSettings),
            new TestCase("worksheet_protection_apis_mutate_expected_settings", WorksheetProtectionApisMutateExpectedSettings),
            new TestCase("autofilter_apis_mutate_expected_settings", AutoFilterApisMutateExpectedSettings),
            new TestCase("defined_name_apis_mutate_expected_settings", DefinedNameApisMutateExpectedSettings),
            new TestCase("hyperlink_apis_mutate_expected_settings", HyperlinkApisMutateExpectedSettings),
            new TestCase("validation_apis_mutate_expected_settings", ValidationApisMutateExpectedSettings),
            new TestCase("conditional_formatting_apis_mutate_expected_settings", ConditionalFormattingApisMutateExpectedSettings),
            new TestCase("conditional_formatting_advanced_apis_mutate_expected_settings", ConditionalFormattingAdvancedApisMutateExpectedSettings),
            new TestCase("page_setup_apis_mutate_expected_settings", PageSetupApisMutateExpectedSettings),
            new TestCase("workbook_metadata_apis_mutate_expected_settings", WorkbookMetadataApisMutateExpectedSettings));
    }

    private static void A1IndexersRoundtrip()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells[2, 27].PutValue("AB3");
        sheet.Cells[0, 0].PutValue(42);

        AssertEx.Equal("AB3", sheet.Cells["AB3"].StringValue);
        AssertEx.Equal("42", sheet.Cells[0, 0].StringValue);
        AssertEx.Equal(CellValueType.String, sheet.Cells["AB3"].Type);
        AssertEx.Equal(CellValueType.Number, sheet.Cells["A1"].Type);
    }

    private static void PutValueOverloadsAssignExpectedTypes()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var timestamp = new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc);

        sheet.Cells["A1"].PutValue("alpha");
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["C1"].PutValue(12.5m);
        sheet.Cells["D1"].PutValue(6.02214076E+23);
        sheet.Cells["E1"].PutValue(true);
        sheet.Cells["F1"].PutValue(timestamp);

        AssertEx.Equal(CellValueType.String, sheet.Cells["A1"].Type);
        AssertEx.Equal(CellValueType.Number, sheet.Cells["B1"].Type);
        AssertEx.Equal(CellValueType.Number, sheet.Cells["C1"].Type);
        AssertEx.Equal(CellValueType.Number, sheet.Cells["D1"].Type);
        AssertEx.Equal(CellValueType.Boolean, sheet.Cells["E1"].Type);
        AssertEx.Equal(CellValueType.DateTime, sheet.Cells["F1"].Type);

        AssertEx.Equal("alpha", sheet.Cells["A1"].Value as string);
        AssertEx.Equal(123, (int)sheet.Cells["B1"].Value!);
        AssertEx.Equal(12.5m, (decimal)sheet.Cells["C1"].Value!);
        AssertEx.True(Math.Abs((double)sheet.Cells["D1"].Value! - 6.02214076E+23) < 1E+10, "Double PutValue should retain magnitude.");
        AssertEx.Equal(true, (bool)sheet.Cells["E1"].Value!);
        AssertEx.Equal(timestamp, (DateTime)sheet.Cells["F1"].Value!);
    }

    private static void StringValueFormatsSupportedScalarTypes()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var timestamp = new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc);

        sheet.Cells["A1"].PutValue(true);
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["C1"].PutValue(12.5m);
        sheet.Cells["D1"].PutValue(timestamp);

        AssertEx.Equal("TRUE", sheet.Cells["A1"].StringValue);
        AssertEx.Equal("TRUE", sheet.Cells["A1"].DisplayStringValue);
        AssertEx.Equal("123", sheet.Cells["B1"].StringValue);
        AssertEx.Equal("123", sheet.Cells["B1"].DisplayStringValue);
        AssertEx.Equal("12.5", sheet.Cells["C1"].StringValue);
        AssertEx.Equal("12.5", sheet.Cells["C1"].DisplayStringValue);
        AssertEx.Equal(timestamp.ToString("M/d/yyyy H:mm", System.Globalization.CultureInfo.InvariantCulture), sheet.Cells["D1"].StringValue);
        AssertEx.Equal(timestamp.ToString("M/d/yyyy H:mm", System.Globalization.CultureInfo.InvariantCulture), sheet.Cells["D1"].DisplayStringValue);
    }

    private static void DisplayStringValueAppliesNumericFormatsAndStringValueStaysRaw()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        var grouped = sheet.Cells["A1"];
        grouped.PutValue(1234.567m);
        var groupedStyle = grouped.GetStyle();
        groupedStyle.NumberFormat = "#,##0.00";
        grouped.SetStyle(groupedStyle);

        var percent = sheet.Cells["B1"];
        percent.PutValue(0.125d);
        var percentStyle = percent.GetStyle();
        percentStyle.NumberFormat = "0.00%";
        percent.SetStyle(percentStyle);

        var scientific = sheet.Cells["C1"];
        scientific.PutValue(1234d);
        var scientificStyle = scientific.GetStyle();
        scientificStyle.NumberFormat = "0.00E+00";
        scientific.SetStyle(scientificStyle);

        var fraction = sheet.Cells["D1"];
        fraction.PutValue(1.25d);
        var fractionStyle = fraction.GetStyle();
        fractionStyle.NumberFormat = "# ?/?";
        fraction.SetStyle(fractionStyle);

        var negative = sheet.Cells["E1"];
        negative.PutValue(-12.3d);
        var negativeStyle = negative.GetStyle();
        negativeStyle.NumberFormat = "#,##0.00_);(#,##0.00)";
        negative.SetStyle(negativeStyle);

        var color = sheet.Cells["F1"];
        color.PutValue(1.25d);
        var colorStyle = color.GetStyle();
        colorStyle.NumberFormat = "[Blue]0.000";
        color.SetStyle(colorStyle);

        var conditionalHigh = sheet.Cells["G1"];
        conditionalHigh.PutValue(125d);
        var conditionalStyle = conditionalHigh.GetStyle();
        conditionalStyle.NumberFormat = "[>100]0.0;\"small\"";
        conditionalHigh.SetStyle(conditionalStyle);

        var conditionalLow = sheet.Cells["H1"];
        conditionalLow.PutValue(10d);
        conditionalLow.SetStyle(conditionalStyle);

        AssertEx.Equal("1234.567", grouped.StringValue);
        AssertEx.Equal("1,234.57", grouped.DisplayStringValue);
        AssertEx.Equal("0.125", percent.StringValue);
        AssertEx.Equal("12.50%", percent.DisplayStringValue);
        AssertEx.Equal("1234", scientific.StringValue);
        AssertEx.Equal("1.23E+03", scientific.DisplayStringValue);
        AssertEx.Equal("1.25", fraction.StringValue);
        AssertEx.Equal("1 1/4", fraction.DisplayStringValue);
        AssertEx.Equal("-12.3", negative.StringValue);
        AssertEx.Equal("(12.30)", negative.DisplayStringValue);
        AssertEx.Equal("1.25", color.StringValue);
        AssertEx.Equal("1.250", color.DisplayStringValue);
        AssertEx.Equal("125", conditionalHigh.StringValue);
        AssertEx.Equal("125.0", conditionalHigh.DisplayStringValue);
        AssertEx.Equal("10", conditionalLow.StringValue);
        AssertEx.Equal("small", conditionalLow.DisplayStringValue);
    }
    private static void DisplayStringValueAppliesDateAndTextFormats()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var timestamp = new DateTime(2024, 5, 6, 7, 8, 9);

        var builtInDate = sheet.Cells["A2"];
        builtInDate.PutValue(timestamp);
        var builtInDateStyle = builtInDate.GetStyle();
        builtInDateStyle.Number = 14;
        builtInDate.SetStyle(builtInDateStyle);

        var customDate = sheet.Cells["B2"];
        customDate.PutValue(timestamp);
        var customDateStyle = customDate.GetStyle();
        customDateStyle.NumberFormat = "d-mmm-yy h:mm AM/PM";
        customDate.SetStyle(customDateStyle);

        var elapsed = sheet.Cells["C2"];
        elapsed.PutValue(timestamp);
        var elapsedStyle = elapsed.GetStyle();
        elapsedStyle.Number = 46;
        elapsed.SetStyle(elapsedStyle);

        var textCell = sheet.Cells["D2"];
        textCell.PutValue("ABC");
        var textStyle = textCell.GetStyle();
        textStyle.NumberFormat = "0;0;0;\"Item \"@";
        textCell.SetStyle(textStyle);

        AssertEx.Equal("5/6/2024 7:08", builtInDate.StringValue);
        AssertEx.Equal("05-06-24", builtInDate.DisplayStringValue);
        AssertEx.Equal("5/6/2024 7:08", customDate.StringValue);
        AssertEx.Equal("6-May-24 7:08 AM", customDate.DisplayStringValue);
        AssertEx.Equal("5/6/2024 7:08", elapsed.StringValue);
        AssertEx.Equal("7:08:09", elapsed.DisplayStringValue);
        AssertEx.Equal("ABC", textCell.StringValue);
        AssertEx.Equal("Item ABC", textCell.DisplayStringValue);
    }
    private static void DisplayStringValueUsesWorkbookCultureAndLocaleDirectives()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var frCulture = CultureInfo.GetCultureInfo("fr-FR");
        var enCulture = CultureInfo.GetCultureInfo("en-US");
        var jpCulture = CultureInfo.GetCultureInfo("ja-JP");
        var deCulture = CultureInfo.GetCultureInfo("de-DE");
        workbook.Settings.Culture = frCulture;

        var numeric = sheet.Cells["A3"];
        numeric.PutValue(1234.5d);
        var numericStyle = numeric.GetStyle();
        numericStyle.NumberFormat = "#,##0.00";
        numeric.SetStyle(numericStyle);

        var dateCell = sheet.Cells["B3"];
        var timestamp = new DateTime(2024, 5, 6, 7, 8, 9);
        dateCell.PutValue(timestamp);
        var dateStyle = dateCell.GetStyle();
        dateStyle.NumberFormat = "dddd, mmmm d, yyyy";
        dateCell.SetStyle(dateStyle);

        var englishDate = sheet.Cells["C3"];
        englishDate.PutValue(timestamp);
        var englishDateStyle = englishDate.GetStyle();
        englishDateStyle.NumberFormat = "[$-409]dddd, mmmm d, yyyy";
        englishDate.SetStyle(englishDateStyle);

        var yenCell = sheet.Cells["D3"];
        yenCell.PutValue(1234.5d);
        var yenStyle = yenCell.GetStyle();
        yenStyle.NumberFormat = "[$\u00A5-411]#,##0.00";
        yenCell.SetStyle(yenStyle);

        var longDateCell = sheet.Cells["E3"];
        longDateCell.PutValue(timestamp);
        var longDateStyle = longDateCell.GetStyle();
        longDateStyle.NumberFormat = "[$-F800]";
        longDateCell.SetStyle(longDateStyle);

        AssertEx.Equal(1234.5d.ToString("#,##0.00", frCulture), numeric.DisplayStringValue);
        AssertEx.Equal(timestamp.ToString("dddd, MMMM d, yyyy", frCulture), dateCell.DisplayStringValue);
        AssertEx.Equal(timestamp.ToString("dddd, MMMM d, yyyy", enCulture), englishDate.DisplayStringValue);
        AssertEx.Equal(1234.5d.ToString("\"\u00A5\"#,##0.00", jpCulture), yenCell.DisplayStringValue);

        workbook.Settings.Culture = deCulture;
        AssertEx.Equal(timestamp.ToString(deCulture.DateTimeFormat.LongDatePattern, deCulture), longDateCell.DisplayStringValue);
    }

    private static void DisplayStringValueAppliesExtendedDateTokens()
    {
        var workbook = new Workbook();
        workbook.Settings.Culture = CultureInfo.GetCultureInfo("en-US");
        var sheet = workbook.Worksheets[0];
        var timestamp = new DateTime(2024, 5, 6, 7, 8, 9, 345);

        var monthInitial = sheet.Cells["A4"];
        monthInitial.PutValue(timestamp);
        var monthInitialStyle = monthInitial.GetStyle();
        monthInitialStyle.NumberFormat = "mmmmm d, yyyy";
        monthInitial.SetStyle(monthInitialStyle);

        var abbreviatedDate = sheet.Cells["B4"];
        abbreviatedDate.PutValue(timestamp);
        var abbreviatedStyle = abbreviatedDate.GetStyle();
        abbreviatedStyle.NumberFormat = "ddd, mmm d yyyy";
        abbreviatedDate.SetStyle(abbreviatedStyle);

        var fractionalSeconds = sheet.Cells["C4"];
        fractionalSeconds.PutValue(timestamp);
        var fractionalStyle = fractionalSeconds.GetStyle();
        fractionalStyle.NumberFormat = "h:mm:ss.000 AM/PM";
        fractionalSeconds.SetStyle(fractionalStyle);

        var shortFraction = sheet.Cells["D4"];
        shortFraction.PutValue(timestamp);
        var shortFractionStyle = shortFraction.GetStyle();
        shortFractionStyle.NumberFormat = "hh:mm:ss.00";
        shortFraction.SetStyle(shortFractionStyle);

        var shortAmPm = sheet.Cells["E4"];
        shortAmPm.PutValue(timestamp);
        var shortAmPmStyle = shortAmPm.GetStyle();
        shortAmPmStyle.NumberFormat = "h A/P";
        shortAmPm.SetStyle(shortAmPmStyle);

        AssertEx.Equal("M 6, 2024", monthInitial.DisplayStringValue);
        AssertEx.Equal("Mon, May 6 2024", abbreviatedDate.DisplayStringValue);
        AssertEx.Equal("7:08:09.345 AM", fractionalSeconds.DisplayStringValue);
        AssertEx.Equal("07:08:09.34", shortFraction.DisplayStringValue);
        AssertEx.Equal("7 A", shortAmPm.DisplayStringValue);
    }

    private static void DisplayStringValueAppliesLongTimeAndElapsedFractionWithCulture()
    {
        var workbook = new Workbook();
        var culture = CultureInfo.GetCultureInfo("de-DE");
        workbook.Settings.Culture = culture;
        var sheet = workbook.Worksheets[0];
        var timestamp = new DateTime(2024, 5, 6, 7, 8, 9, 345);

        var longTimeCell = sheet.Cells["F4"];
        longTimeCell.PutValue(timestamp);
        var longTimeStyle = longTimeCell.GetStyle();
        longTimeStyle.NumberFormat = "[$-F400]";
        longTimeCell.SetStyle(longTimeStyle);

        var elapsedCell = sheet.Cells["G4"];
        elapsedCell.PutValue(timestamp);
        var elapsedStyle = elapsedCell.GetStyle();
        elapsedStyle.NumberFormat = "[h]:mm:ss.000";
        elapsedCell.SetStyle(elapsedStyle);

        AssertEx.Equal(timestamp.ToString(culture.DateTimeFormat.LongTimePattern, culture), longTimeCell.DisplayStringValue);
        AssertEx.Equal("7:08:09" + culture.NumberFormat.NumberDecimalSeparator + "345", elapsedCell.DisplayStringValue);
    }

    private static void FormulaPropertyNormalizesAndPreservesCachedValue()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["C3"];

        cell.PutValue(20);
        cell.Formula = "A1+B1";

        AssertEx.Equal(CellValueType.Formula, cell.Type);
        AssertEx.Equal("=A1+B1", cell.Formula);
        AssertEx.Equal("20", cell.StringValue);
        AssertEx.Equal(20, (int)cell.Value!);
    }

    private static void BlankCellsAreBlankByDefault()
    {
        var workbook = new Workbook();
        var cell = workbook.Worksheets[0].Cells["Z99"];

        AssertEx.Equal(CellValueType.Blank, cell.Type);
        AssertEx.Null(cell.Value);
        AssertEx.Equal(string.Empty, cell.StringValue);
    }

    private static void WorksheetNameAndCollectionGuards()
    {
        var workbook = new Workbook();
        workbook.Worksheets.Add("Data");

        AssertEx.Throws<CellsException>(delegate { workbook.Worksheets.Add("data"); });
        AssertEx.Throws<CellsException>(delegate { _ = workbook.Worksheets[0].Cells["1A"]; });
        AssertEx.Throws<CellsException>(delegate { _ = workbook.Worksheets[0].Cells[-1, 0]; });

        var parsed = CellAddress.Parse("AB3");
        AssertEx.Equal(2, parsed.RowIndex);
        AssertEx.Equal(27, parsed.ColumnIndex);
        AssertEx.Equal("AB3", parsed.ToString());
    }

    private static void StyleMutationRequiresSetStyleAndReturnsClones()
    {
        var workbook = new Workbook();
        var cell = workbook.Worksheets[0].Cells["A1"];

        var style = cell.GetStyle();
        style.Font.Bold = true;
        style.HorizontalAlignment = HorizontalAlignmentType.Right;

        var untouched = cell.GetStyle();
        AssertEx.False(untouched.Font.Bold);
        AssertEx.Equal(HorizontalAlignmentType.General, untouched.HorizontalAlignment);

        cell.SetStyle(style);
        var applied = cell.GetStyle();
        AssertEx.True(applied.Font.Bold);
        AssertEx.Equal(HorizontalAlignmentType.Right, applied.HorizontalAlignment);

        applied.Font.Italic = true;
        AssertEx.False(cell.GetStyle().Font.Italic);
    }

    private static void StyleApiCoversAllPublicSettings()
    {
        var workbook = new Workbook();
        var primaryCell = workbook.Worksheets[0].Cells["A1"];
        primaryCell.PutValue(1);

        var primaryStyle = primaryCell.GetStyle();
        StyleScenarioFactory.ApplyPrimaryStyle(primaryStyle);

        var untouched = primaryCell.GetStyle();
        AssertEx.Equal("Calibri", untouched.Font.Name);
        AssertEx.Equal(FillPattern.None, untouched.Pattern);
        AssertEx.True(untouched.IsLocked);
        AssertEx.False(untouched.IsHidden);
        AssertEx.Equal(0, untouched.IndentLevel);
        AssertEx.Equal(0, untouched.TextRotation);
        AssertEx.Equal(0, untouched.ReadingOrder);
        AssertEx.False(untouched.ShrinkToFit);
        AssertEx.False(untouched.Font.StrikeThrough);
        AssertEx.False(untouched.Borders.DiagonalUp);
        AssertEx.False(untouched.Borders.DiagonalDown);
        AssertEx.Equal("General", untouched.NumberFormat);

        primaryCell.SetStyle(primaryStyle);
        StyleScenarioFactory.AssertPrimaryStyle(primaryCell.GetStyle());

        var mutatedClone = primaryCell.GetStyle();
        mutatedClone.Font.Name = "Mutated";
        mutatedClone.Font.StrikeThrough = false;
        mutatedClone.Pattern = FillPattern.None;
        mutatedClone.ForegroundColor = Color.Empty;
        mutatedClone.BackgroundColor = Color.Empty;
        mutatedClone.Borders.Right.LineStyle = BorderStyleType.None;
        mutatedClone.Borders.Diagonal.LineStyle = BorderStyleType.None;
        mutatedClone.Borders.DiagonalUp = false;
        mutatedClone.Borders.DiagonalDown = false;
        mutatedClone.Number = 0;
        mutatedClone.HorizontalAlignment = HorizontalAlignmentType.Left;
        mutatedClone.VerticalAlignment = VerticalAlignmentType.Bottom;
        mutatedClone.WrapText = false;
        mutatedClone.IndentLevel = 0;
        mutatedClone.TextRotation = 0;
        mutatedClone.ShrinkToFit = false;
        mutatedClone.ReadingOrder = 0;
        mutatedClone.RelativeIndent = 0;
        mutatedClone.IsLocked = true;
        mutatedClone.IsHidden = false;

        var persistedPrimary = primaryCell.GetStyle();
        AssertEx.Equal("Arial", persistedPrimary.Font.Name);
        AssertEx.True(persistedPrimary.Font.StrikeThrough);
        AssertEx.Equal(FillPattern.LightGrid, persistedPrimary.Pattern);
        AssertEx.Equal(BorderStyleType.MediumDashDot, persistedPrimary.Borders.Right.LineStyle);
        AssertEx.Equal(BorderStyleType.SlantedDashDot, persistedPrimary.Borders.Diagonal.LineStyle);
        AssertEx.True(persistedPrimary.Borders.DiagonalUp);
        AssertEx.True(persistedPrimary.Borders.DiagonalDown);
        AssertEx.Equal(4, persistedPrimary.Number);
        AssertEx.Equal("#,##0.00", persistedPrimary.NumberFormat);
        AssertEx.Equal(HorizontalAlignmentType.Distributed, persistedPrimary.HorizontalAlignment);
        AssertEx.Equal(VerticalAlignmentType.Distributed, persistedPrimary.VerticalAlignment);
        AssertEx.True(persistedPrimary.WrapText);
        AssertEx.Equal(2, persistedPrimary.IndentLevel);
        AssertEx.Equal(45, persistedPrimary.TextRotation);
        AssertEx.True(persistedPrimary.ShrinkToFit);
        AssertEx.Equal(2, persistedPrimary.ReadingOrder);
        AssertEx.Equal(1, persistedPrimary.RelativeIndent);
        AssertEx.False(persistedPrimary.IsLocked);
        AssertEx.True(persistedPrimary.IsHidden);

        var customCell = workbook.Worksheets[0].Cells["B2"];
        var customStyle = customCell.GetStyle();
        StyleScenarioFactory.ApplyCustomNumberStyle(customStyle);
        customCell.SetStyle(customStyle);

        StyleScenarioFactory.AssertCustomNumberStyle(customCell.GetStyle());
        AssertEx.Equal(CellValueType.Blank, customCell.Type);

        var numberFormatStyle = new Style();
        numberFormatStyle.NumberFormat = "0.00%";
        AssertEx.Equal(10, numberFormatStyle.Number);
        AssertEx.Null(numberFormatStyle.Custom);
        numberFormatStyle.NumberFormat = "[Blue]0.000";
        AssertEx.Equal(0, numberFormatStyle.Number);
        AssertEx.Equal("[Blue]0.000", numberFormatStyle.Custom);

        AssertEx.Throws<CellsException>(delegate { numberFormatStyle.IndentLevel = -1; });
        AssertEx.Throws<CellsException>(delegate { numberFormatStyle.TextRotation = 181; });
        AssertEx.Throws<CellsException>(delegate { numberFormatStyle.ReadingOrder = 3; });
    }
    private static void WorksheetRowColumnAndMergeApisMutateExpectedSettings()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        sheet.VisibilityType = VisibilityType.Hidden;
        sheet.Cells.Rows[2].Height = 19.75d;
        sheet.Cells.Rows[4].IsHidden = true;
        sheet.Cells.Columns[1].Width = 25.5d;
        sheet.Cells.Columns[3].IsHidden = true;
        sheet.Cells.Merge(1, 1, 2, 3);

        AssertEx.Equal(VisibilityType.Hidden, sheet.VisibilityType);
        AssertEx.Equal(19.75d, sheet.Cells.Rows[2].Height ?? 0d);
        AssertEx.True(sheet.Cells.Rows[4].IsHidden);
        AssertEx.Equal(25.5d, sheet.Cells.Columns[1].Width ?? 0d);
        AssertEx.True(sheet.Cells.Columns[3].IsHidden);
        AssertEx.Equal(1, sheet.Cells.MergedCells.Count);
        AssertEx.Equal(1, sheet.Cells.MergedCells[0].FirstRow);
        AssertEx.Equal(1, sheet.Cells.MergedCells[0].FirstColumn);
        AssertEx.Equal(2, sheet.Cells.MergedCells[0].TotalRows);
        AssertEx.Equal(3, sheet.Cells.MergedCells[0].TotalColumns);
        AssertEx.Throws<CellsException>(delegate { sheet.Cells.Merge(2, 2, 2, 2); });
    }

    private static void WorksheetViewApisMutateExpectedSettings()
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

        sheet.TabColor = Color.Empty;
        AssertEx.Equal(Color.Empty, sheet.TabColor);
        AssertEx.Throws<CellsException>(delegate { sheet.Zoom = 9; });
        AssertEx.Throws<CellsException>(delegate { sheet.Zoom = 401; });
    }

    private static void WorksheetProtectionApisMutateExpectedSettings()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        sheet.Protect();
        sheet.Protection.Objects = true;
        sheet.Protection.FormatCells = true;
        sheet.Protection.InsertRows = true;
        sheet.Protection.SelectUnlockedCells = true;

        AssertEx.True(sheet.Protection.IsProtected);
        AssertEx.True(sheet.Protection.Objects);
        AssertEx.True(sheet.Protection.FormatCells);
        AssertEx.True(sheet.Protection.InsertRows);
        AssertEx.True(sheet.Protection.SelectUnlockedCells);

        sheet.Unprotect();
        AssertEx.False(sheet.Protection.IsProtected);
        AssertEx.False(sheet.Protection.Objects);
        AssertEx.False(sheet.Protection.FormatCells);
        AssertEx.False(sheet.Protection.InsertRows);
        AssertEx.False(sheet.Protection.SelectUnlockedCells);

        sheet.Protection.AutoFilter = true;
        AssertEx.True(sheet.Protection.IsProtected);
        AssertEx.True(sheet.Protection.AutoFilter);
    }

    private static void AutoFilterApisMutateExpectedSettings()
    {
        var workbook = AutoFilterScenarioFactory.CreateAutoFilterWorkbook();
        AutoFilterScenarioFactory.AssertAutoFilter(workbook);

        var sheet = workbook.Worksheets[0];
        AssertEx.Throws<CellsException>(delegate { sheet.AutoFilter.FilterColumns.Add(-1); });
        AssertEx.Throws<CellsException>(delegate { sheet.AutoFilter.FilterColumns.Add(0); });
        AssertEx.Throws<CellsException>(delegate { sheet.AutoFilter.SortState.SortConditions.Add("1A"); });

        sheet.AutoFilter.FilterColumns.RemoveAt(4);
        AssertEx.Equal(4, sheet.AutoFilter.FilterColumns.Count);
        sheet.AutoFilter.Clear();
        AssertEx.Equal(string.Empty, sheet.AutoFilter.Range);
        AssertEx.Equal(0, sheet.AutoFilter.FilterColumns.Count);
        AssertEx.Equal(0, sheet.AutoFilter.SortState.SortConditions.Count);
    }

    private static void DefinedNameApisMutateExpectedSettings()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Name = "Data";
        workbook.Worksheets.Add("Scoped");

        var total = workbook.DefinedNames[workbook.DefinedNames.Add("Total", "=SUM(Data!$A$1:$A$2)")];
        total.Hidden = true;
        total.Comment = "Workbook scope";

        var scoped = workbook.DefinedNames[workbook.DefinedNames.Add("Input", "'Scoped'!$B$2", 1)];
        scoped.Comment = "Local scope";

        AssertEx.Equal(2, workbook.DefinedNames.Count);
        AssertEx.Equal("Total", total.Name);
        AssertEx.Equal("SUM(Data!$A$1:$A$2)", total.Formula);
        AssertEx.Null(total.LocalSheetIndex);
        AssertEx.True(total.Hidden);
        AssertEx.Equal("Workbook scope", total.Comment);

        AssertEx.Equal("Input", scoped.Name);
        AssertEx.Equal("'Scoped'!$B$2", scoped.Formula);
        AssertEx.Equal(1, scoped.LocalSheetIndex ?? -1);
        AssertEx.Equal("Local scope", scoped.Comment);

        AssertEx.Throws<CellsException>(delegate { workbook.DefinedNames.Add("Total", "1"); });
        AssertEx.Throws<CellsException>(delegate { workbook.DefinedNames.Add("_xlnm.Print_Area", "A1"); });
        AssertEx.Throws<CellsException>(delegate { workbook.DefinedNames.Add("Broken", "1", 5); });

        scoped.Name = "Total";
        AssertEx.Throws<CellsException>(delegate { scoped.LocalSheetIndex = null; });

        workbook.DefinedNames.RemoveAt(1);
        AssertEx.Equal(1, workbook.DefinedNames.Count);
        AssertEx.Throws<CellsException>(delegate { workbook.DefinedNames.RemoveAt(5); });
    }

    private static void HyperlinkApisMutateExpectedSettings()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        sheet.Cells["A1"].PutValue("Docs");
        var external = sheet.Hyperlinks[sheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/docs")];
        external.TextToDisplay = "Docs";
        external.ScreenTip = "External docs";

        var internalLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("B2", 1, 1, "Sheet1!C3")];
        internalLink.TextToDisplay = "Jump";

        var rangeLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("C4", 2, 2, "mailto:test@example.com")];
        rangeLink.ScreenTip = "Send mail";

        AssertEx.Equal(3, sheet.Hyperlinks.Count);
        AssertEx.Equal("A1", external.Area);
        AssertEx.Equal("https://example.com/docs", external.Address);
        AssertEx.Equal("External docs", external.ScreenTip);
        AssertEx.Equal("Docs", external.TextToDisplay);

        AssertEx.Equal("B2", internalLink.Area);
        AssertEx.Equal("Sheet1!C3", internalLink.Address);
        AssertEx.Equal("Jump", internalLink.TextToDisplay);

        AssertEx.Equal("C4:D5", rangeLink.Area);
        AssertEx.Equal("mailto:test@example.com", rangeLink.Address);
        AssertEx.Equal("Send mail", rangeLink.ScreenTip);

        AssertEx.Throws<CellsException>(delegate { sheet.Hyperlinks.Add("A1", 1, 1, "https://overlap.example.com"); });
        AssertEx.Throws<CellsException>(delegate { sheet.Hyperlinks.Add("Z1", 0, 1, "https://invalid.example.com"); });
        AssertEx.Throws<CellsException>(delegate { sheet.Hyperlinks.Add("A2", 1, 1, string.Empty); });
        AssertEx.Throws<CellsException>(delegate { _ = sheet.Hyperlinks[-1]; });

        sheet.Hyperlinks.RemoveAt(1);
        AssertEx.Equal(2, sheet.Hyperlinks.Count);
        AssertEx.Equal("C4:D5", sheet.Hyperlinks[1].Area);
        AssertEx.Throws<CellsException>(delegate { sheet.Hyperlinks.RemoveAt(99); });
    }

    private static void ValidationApisMutateExpectedSettings()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        var primary = sheet.Validations[sheet.Validations.Add(CellArea.CreateCellArea("A1", "B2"))];
        primary.Type = ValidationType.WholeNumber;
        primary.Operator = OperatorType.Between;
        primary.Formula1 = "=1";
        primary.Formula2 = "=10";
        primary.ShowError = true;
        primary.ErrorTitle = "Whole Number";
        primary.ErrorMessage = "Enter 1-10";
        primary.AddArea(CellArea.CreateCellArea("D4", "D5"));

        AssertEx.Equal(1, sheet.Validations.Count);
        AssertEx.Equal(2, primary.Areas.Count);
        AssertEx.Equal("1", primary.Formula1);
        AssertEx.Equal("10", primary.Formula2);
        AssertEx.Equal(ValidationType.WholeNumber, sheet.Validations.GetValidationInCell(0, 0)!.Type);
        AssertEx.Equal(ValidationType.WholeNumber, sheet.Validations.GetValidationInCell(4, 3)!.Type);
        AssertEx.Throws<CellsException>(delegate { sheet.Validations.Add(CellArea.CreateCellArea("B2", "C3")); });

        var secondIndex = sheet.Validations.Add(CellArea.CreateCellArea("F1", "F1"));
        var second = sheet.Validations[secondIndex];
        second.Type = ValidationType.List;
        second.Formula1 = "\"Y,N\"";
        AssertEx.Equal(2, sheet.Validations.Count);

        sheet.Validations.RemoveACell(0, 0);
        AssertEx.Null(sheet.Validations.GetValidationInCell(0, 0));
        AssertEx.NotNull(sheet.Validations.GetValidationInCell(0, 1));

        sheet.Validations.RemoveArea(CellArea.CreateCellArea("F1", "F1"));
        AssertEx.Equal(1, sheet.Validations.Count);
        AssertEx.Throws<CellsException>(delegate { sheet.Validations.RemoveACell(-1, 0); });
        AssertEx.Throws<CellsException>(delegate { _ = sheet.Validations[-1]; });
    }
    private static void ConditionalFormattingApisMutateExpectedSettings()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        var index = sheet.ConditionalFormattings.Add();
        var collection = sheet.ConditionalFormattings[index];
        collection.AddArea(CellArea.CreateCellArea("A1", "A3"));
        var conditionIndex = collection.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "=1", "=9");
        var condition = collection[conditionIndex];
        condition.StopIfTrue = true;
        condition.Priority = 1;
        var style = condition.Style;
        style.Pattern = FillPattern.Solid;
        style.ForegroundColor = Color.FromArgb(255, 255, 0, 0);
        condition.Style = style;

        AssertEx.Equal(1, sheet.ConditionalFormattings.Count);
        AssertEx.Equal(1, collection.RangeCount);
        AssertEx.Equal(1, collection.Count);
        AssertEx.Equal("1", condition.Formula1);
        AssertEx.Equal("9", condition.Formula2);
        AssertEx.True(condition.StopIfTrue);
        AssertEx.Equal(FillPattern.Solid, condition.Style.Pattern);

        collection.AddCondition(FormatConditionType.Expression);
        AssertEx.Equal(2, collection.Count);
        collection.RemoveCondition(1);
        AssertEx.Equal(1, collection.Count);

        collection.AddArea(CellArea.CreateCellArea("C1", "C2"));
        AssertEx.Equal(2, collection.RangeCount);
        collection.RemoveArea(0, 0, 1, 1);
        AssertEx.Equal(2, collection.RangeCount);
        AssertEx.Equal(0, collection.GetCellArea(0).FirstRow);
        AssertEx.Equal(2, collection.GetCellArea(0).FirstColumn);

        sheet.ConditionalFormattings.RemoveArea(0, 2, 2, 1);
        AssertEx.Equal(1, sheet.ConditionalFormattings.Count);
        sheet.ConditionalFormattings.RemoveAt(0);
        AssertEx.Equal(0, sheet.ConditionalFormattings.Count);
        AssertEx.Throws<CellsException>(delegate { _ = sheet.ConditionalFormattings[-1]; });
    }
    private static void ConditionalFormattingAdvancedApisMutateExpectedSettings()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        var contains = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        contains.AddArea(CellArea.CreateCellArea("A1", "A10"));
        var containsRule = contains[contains.AddCondition(FormatConditionType.ContainsText)];
        containsRule.Formula1 = "error";
        containsRule.Priority = 2;

        var timePeriod = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        timePeriod.AddArea(CellArea.CreateCellArea("B1", "B10"));
        var timeRule = timePeriod[timePeriod.AddCondition(FormatConditionType.TimePeriod)];
        timeRule.TimePeriod = "today";

        var top10 = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        top10.AddArea(CellArea.CreateCellArea("C1", "C10"));
        var topRule = top10[top10.AddCondition(FormatConditionType.Top10)];
        topRule.Percent = true;
        topRule.Rank = 10;

        var colorScale = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        colorScale.AddArea(CellArea.CreateCellArea("D1", "D10"));
        var colorRule = colorScale[colorScale.AddCondition(FormatConditionType.ColorScale)];
        colorRule.ColorScaleCount = 3;
        colorRule.MinColor = Color.FromArgb(255, 248, 105, 107);
        colorRule.MidColor = Color.FromArgb(255, 255, 235, 132);
        colorRule.MaxColor = Color.FromArgb(255, 99, 190, 123);

        var dataBar = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        dataBar.AddArea(CellArea.CreateCellArea("E1", "E10"));
        var dataBarRule = dataBar[dataBar.AddCondition(FormatConditionType.DataBar)];
        dataBarRule.BarColor = Color.FromArgb(255, 99, 142, 198);
        dataBarRule.ShowBorder = true;
        dataBarRule.Direction = "left-to-right";

        var iconSet = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        iconSet.AddArea(CellArea.CreateCellArea("F1", "F10"));
        var iconSetRule = iconSet[iconSet.AddCondition(FormatConditionType.IconSet)];
        iconSetRule.IconSetType = "4Arrows";
        iconSetRule.ReverseIcons = true;
        iconSetRule.ShowIconOnly = true;

        AssertEx.Equal(6, sheet.ConditionalFormattings.Count);
        AssertEx.Equal(FormatConditionType.ContainsText, containsRule.Type);
        AssertEx.Equal("error", containsRule.Formula1);
        AssertEx.Equal(2, containsRule.Priority);
        AssertEx.Equal("today", timeRule.TimePeriod);
        AssertEx.True(topRule.Percent);
        AssertEx.Equal(10, topRule.Rank);
        AssertEx.Equal(3, colorRule.ColorScaleCount);
        AssertEx.Equal(Color.FromArgb(255, 248, 105, 107), colorRule.MinColor);
        AssertEx.Equal(Color.FromArgb(255, 99, 142, 198), dataBarRule.BarColor);
        AssertEx.True(dataBarRule.ShowBorder);
        AssertEx.Equal("4Arrows", iconSetRule.IconSetType);
        AssertEx.True(iconSetRule.ReverseIcons);
        AssertEx.True(iconSetRule.ShowIconOnly);
    }
    private static void PageSetupApisMutateExpectedSettings()
    {
        var workbook = new Workbook();
        var pageSetup = workbook.Worksheets[0].PageSetup;

        pageSetup.LeftMargin = 0.508d;
        pageSetup.RightMargin = 0.635d;
        pageSetup.Orientation = PageOrientationType.Landscape;
        pageSetup.PaperSize = PaperSizeType.PaperA4;
        pageSetup.FirstPageNumber = 2;
        pageSetup.Scale = 90;
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = 3;
        pageSetup.PrintArea = "$A$1:$D$20";
        pageSetup.PrintTitleRows = "$1:$2";
        pageSetup.PrintTitleColumns = "$A:$B";
        pageSetup.LeftHeader = "LH";
        pageSetup.CenterFooter = "CF";
        pageSetup.PrintGridlines = true;
        pageSetup.CenterHorizontally = true;
        pageSetup.AddHorizontalPageBreak(5);
        pageSetup.AddVerticalPageBreak(2);

        AssertEx.Equal(0.508d, pageSetup.LeftMargin);
        AssertEx.Equal(0.635d, pageSetup.RightMargin);
        AssertEx.Equal(PageOrientationType.Landscape, pageSetup.Orientation);
        AssertEx.Equal(PaperSizeType.PaperA4, pageSetup.PaperSize);
        AssertEx.Equal(2, pageSetup.FirstPageNumber ?? 0);
        AssertEx.Equal(90, pageSetup.Scale ?? 0);
        AssertEx.Equal(1, pageSetup.FitToPagesWide ?? 0);
        AssertEx.Equal(3, pageSetup.FitToPagesTall ?? 0);
        AssertEx.Equal("$A$1:$D$20", pageSetup.PrintArea);
        AssertEx.Equal("$1:$2", pageSetup.PrintTitleRows);
        AssertEx.Equal("$A:$B", pageSetup.PrintTitleColumns);
        AssertEx.Equal("LH", pageSetup.LeftHeader);
        AssertEx.Equal("CF", pageSetup.CenterFooter);
        AssertEx.True(pageSetup.PrintGridlines);
        AssertEx.True(pageSetup.CenterHorizontally);
        AssertEx.Equal(1, pageSetup.HorizontalPageBreaks.Count);
        AssertEx.Equal(5, pageSetup.HorizontalPageBreaks[0]);
        AssertEx.Equal(1, pageSetup.VerticalPageBreaks.Count);
        AssertEx.Equal(2, pageSetup.VerticalPageBreaks[0]);
        AssertEx.Throws<CellsException>(delegate { pageSetup.Scale = 5; });
        AssertEx.Throws<CellsException>(delegate { pageSetup.LeftMargin = -0.1d; });
    }

    private static void WorkbookMetadataApisMutateExpectedSettings()
    {
        var workbook = WorkbookMetadataScenarioFactory.CreateWorkbookMetadataWorkbook();
        WorkbookMetadataScenarioFactory.AssertWorkbookMetadata(workbook);
    }}





















