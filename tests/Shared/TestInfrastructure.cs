using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Testing;

public sealed class TestCase
{
    public TestCase(string name, Action body)
    {
        Name = name;
        Body = body;
    }

    public string Name { get; }
    public Action Body { get; }
}

public static class TestRunner
{
    public static int Run(string suiteName, params TestCase[] tests)
    {
        Console.WriteLine($"Running {suiteName} ({tests.Length} tests)");
        var failed = 0;

        foreach (var test in tests)
        {
            try
            {
                test.Body();
                Console.WriteLine($"PASS {test.Name}");
            }
            catch (Exception exception)
            {
                failed++;
                Console.WriteLine($"FAIL {test.Name}");
                Console.WriteLine(exception);
            }
        }

        Console.WriteLine($"{suiteName}: {tests.Length - failed} passed, {failed} failed");
        return failed == 0 ? 0 : 1;
    }
}

public static class AssertEx
{
    public static void True(bool condition, string? message = null)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message ?? "Expected condition to be true.");
        }
    }

    public static void False(bool condition, string? message = null)
    {
        if (condition)
        {
            throw new InvalidOperationException(message ?? "Expected condition to be false.");
        }
    }

    public static void Equal<T>(T expected, T actual, string? message = null)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException(message ?? $"Expected '{expected}', got '{actual}'.");
        }
    }

    public static void NotNull(object? value, string? message = null)
    {
        if (value is null)
        {
            throw new InvalidOperationException(message ?? "Expected value to be non-null.");
        }
    }

    public static void Null(object? value, string? message = null)
    {
        if (value is not null)
        {
            throw new InvalidOperationException(message ?? $"Expected value to be null, got '{value}'.");
        }
    }

    public static void Contains(string expectedSubstring, string actual, string? message = null)
    {
        if (actual.IndexOf(expectedSubstring, StringComparison.Ordinal) < 0)
        {
            throw new InvalidOperationException(message ?? $"Expected '{actual}' to contain '{expectedSubstring}'.");
        }
    }

    public static T Throws<T>(Action action, string? message = null) where T : Exception
    {
        try
        {
            action();
        }
        catch (T exception)
        {
            return exception;
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException(message ?? $"Expected {typeof(T).Name}, got {exception.GetType().Name}.", exception);
        }

        throw new InvalidOperationException(message ?? $"Expected {typeof(T).Name} to be thrown.");
    }
}

public sealed class TemporaryDirectory : IDisposable
{
    public TemporaryDirectory(string suiteName)
    {
        var repositoryRoot = ResolveRepositoryRoot();
        RootPath = Path.Combine(repositoryRoot, "output", suiteName, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(RootPath);
    }

    public string RootPath { get; }

    public string GetPath(string fileName)
    {
        return Path.Combine(RootPath, fileName);
    }

    public void Dispose()
    {
        // Preserve generated artifacts under the repo-local output folder for inspection.
    }

    private static string ResolveRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Aspose.Cells_FOSS.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        return Directory.GetCurrentDirectory();
    }
}

public static class WorkbookScenarioFactory
{
    public static Workbook CreateMixedCellWorkbook(bool useDate1904 = false)
    {
        var workbook = new Workbook();
        workbook.Settings.Date1904 = useDate1904;

        var sheet = workbook.Worksheets[0];
        sheet.Name = "Data";
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["C1"].PutValue(true);
        sheet.Cells["D1"].PutValue(12.5m);
        sheet.Cells["E1"].PutValue(6.02214076E+23);
        sheet.Cells["F1"].PutValue(new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc));
        sheet.Cells["G1"].PutValue(20);
        sheet.Cells["G1"].Formula = "=B1*2";
        return workbook;
    }
}

public static class WorksheetScenarioFactory
{
    public static Workbook CreateWorksheetSettingsWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Layout";
        sheet.VisibilityType = VisibilityType.Hidden;
        sheet.TabColor = Color.FromArgb(255, 34, 68, 102);
        sheet.ShowGridlines = false;
        sheet.ShowRowColumnHeaders = false;
        sheet.ShowZeros = false;
        sheet.RightToLeft = true;
        sheet.Zoom = 85;
        sheet.Protect();
        sheet.Protection.Objects = true;
        sheet.Protection.Scenarios = true;
        sheet.Protection.FormatCells = true;
        sheet.Protection.InsertRows = true;
        sheet.Protection.AutoFilter = true;
        sheet.Protection.SelectLockedCells = true;
        sheet.Protection.SelectUnlockedCells = true;

        sheet.Cells["A1"].PutValue("Merged");
        sheet.Cells["C4"].PutValue(99);
        sheet.Cells.Rows[1].Height = 22.5d;
        sheet.Cells.Rows[3].IsHidden = true;
        sheet.Cells.Columns[0].Width = 18.25d;
        sheet.Cells.Columns[2].IsHidden = true;
        sheet.Cells.Merge(0, 0, 2, 2);

        var visibleSheetIndex = workbook.Worksheets.Add();
        var visibleSheet = workbook.Worksheets[visibleSheetIndex];
        visibleSheet.Name = "Visible";
        visibleSheet.Cells["A1"].PutValue("Visible");
        workbook.Worksheets.ActiveSheetName = "Visible";

        return workbook;
    }

    public static void AssertWorksheetSettings(Workbook workbook)
    {
        var sheet = workbook.Worksheets["Layout"];
        AssertEx.NotNull(sheet, "Expected the worksheet settings scenario to contain the 'Layout' sheet.");
        AssertEx.Equal(VisibilityType.Hidden, sheet!.VisibilityType);
        AssertEx.Equal(Color.FromArgb(255, 34, 68, 102), sheet.TabColor);
        AssertEx.False(sheet.ShowGridlines);
        AssertEx.False(sheet.ShowRowColumnHeaders);
        AssertEx.False(sheet.ShowZeros);
        AssertEx.True(sheet.RightToLeft);
        AssertEx.Equal(85, sheet.Zoom);
        AssertEx.True(sheet.Protection.IsProtected);
        AssertEx.True(sheet.Protection.Objects);
        AssertEx.True(sheet.Protection.Scenarios);
        AssertEx.True(sheet.Protection.FormatCells);
        AssertEx.True(sheet.Protection.InsertRows);
        AssertEx.True(sheet.Protection.AutoFilter);
        AssertEx.True(sheet.Protection.SelectLockedCells);
        AssertEx.True(sheet.Protection.SelectUnlockedCells);
        AssertEx.Equal("Merged", sheet.Cells["A1"].StringValue);
        AssertEx.Equal(99, (int)sheet.Cells["C4"].Value!);
        AssertEx.Equal(22.5d, sheet.Cells.Rows[1].Height ?? 0d);
        AssertEx.True(sheet.Cells.Rows[3].IsHidden);
        AssertEx.Equal(18.25d, sheet.Cells.Columns[0].Width ?? 0d);
        AssertEx.True(sheet.Cells.Columns[2].IsHidden);
        AssertEx.Equal(1, sheet.Cells.MergedCells.Count);
        AssertEx.Equal(0, sheet.Cells.MergedCells[0].FirstRow);
        AssertEx.Equal(0, sheet.Cells.MergedCells[0].FirstColumn);
        AssertEx.Equal(2, sheet.Cells.MergedCells[0].TotalRows);
        AssertEx.Equal(2, sheet.Cells.MergedCells[0].TotalColumns);
    }

    public static void AssertWorksheetSettingsScenarioHasVisibleSheet(Workbook workbook)
    {
        AssertEx.True(workbook.Worksheets.Count >= 2);
        var visibleSheet = workbook.Worksheets["Visible"];
        AssertEx.NotNull(visibleSheet, "Expected the worksheet settings scenario to contain a visible sheet.");
        AssertEx.Equal(VisibilityType.Visible, visibleSheet!.VisibilityType);
        AssertEx.Equal("Visible", visibleSheet.Cells["A1"].StringValue);
        AssertEx.Equal("Visible", workbook.Worksheets.ActiveSheetName);
    }
}
public static class PageSetupScenarioFactory
{
    public static Workbook CreatePageSetupWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Print Sheet";
        sheet.Cells["A1"].PutValue("Title");
        sheet.Cells["C10"].PutValue(42);

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
        pageSetup.PrintArea = "$A$1:$C$10";
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

        return workbook;
    }

    public static void AssertPageSetup(Workbook workbook)
    {
        var pageSetup = workbook.Worksheets[0].PageSetup;
        AssertEx.Equal(0.25d, pageSetup.LeftMarginInch);
        AssertEx.Equal(0.4d, pageSetup.RightMarginInch);
        AssertEx.Equal(0.5d, pageSetup.TopMarginInch);
        AssertEx.Equal(0.6d, pageSetup.BottomMarginInch);
        AssertEx.Equal(0.2d, pageSetup.HeaderMarginInch);
        AssertEx.Equal(0.22d, pageSetup.FooterMarginInch);
        AssertEx.Equal(PageOrientationType.Landscape, pageSetup.Orientation);
        AssertEx.Equal(PaperSizeType.PaperA4, pageSetup.PaperSize);
        AssertEx.Equal(3, pageSetup.FirstPageNumber ?? 0);
        AssertEx.Equal(95, pageSetup.Scale ?? 0);
        AssertEx.Equal(1, pageSetup.FitToPagesWide ?? 0);
        AssertEx.Equal(2, pageSetup.FitToPagesTall ?? 0);
        AssertEx.Equal("$A$1:$C$10", pageSetup.PrintArea);
        AssertEx.Equal("$1:$2", pageSetup.PrintTitleRows);
        AssertEx.Equal("$A:$B", pageSetup.PrintTitleColumns);
        AssertEx.Equal("Left Header", pageSetup.LeftHeader);
        AssertEx.Equal("Center Header", pageSetup.CenterHeader);
        AssertEx.Equal("Right Header", pageSetup.RightHeader);
        AssertEx.Equal("Left Footer", pageSetup.LeftFooter);
        AssertEx.Equal("Center Footer", pageSetup.CenterFooter);
        AssertEx.Equal("Right Footer", pageSetup.RightFooter);
        AssertEx.True(pageSetup.PrintGridlines);
        AssertEx.True(pageSetup.PrintHeadings);
        AssertEx.True(pageSetup.CenterHorizontally);
        AssertEx.True(pageSetup.CenterVertically);
        AssertEx.Equal(2, pageSetup.HorizontalPageBreaks.Count);
        AssertEx.Equal(4, pageSetup.HorizontalPageBreaks[0]);
        AssertEx.Equal(7, pageSetup.HorizontalPageBreaks[1]);
        AssertEx.Equal(1, pageSetup.VerticalPageBreaks.Count);
        AssertEx.Equal(2, pageSetup.VerticalPageBreaks[0]);
        AssertEx.Equal("Title", workbook.Worksheets[0].Cells["A1"].StringValue);
        AssertEx.Equal("42", workbook.Worksheets[0].Cells["C10"].StringValue);
    }
}

public static class HyperlinkScenarioFactory
{
    public static Workbook CreateHyperlinkWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Links";
        workbook.Worksheets.Add("Target Sheet");
        workbook.Worksheets[1].Cells[2, 2].PutValue("Target");

        sheet.Cells["A1"].PutValue("Docs");
        var external = sheet.Hyperlinks[sheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/docs?q=1")];
        external.TextToDisplay = "Docs";
        external.ScreenTip = "External docs";

        sheet.Cells["B2"].PutValue("Jump");
        var internalLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("B2", 1, 1, "'Target Sheet'!C3")];
        internalLink.TextToDisplay = "Jump";
        internalLink.ScreenTip = "Jump to target";

        var rangeLink = sheet.Hyperlinks[sheet.Hyperlinks.Add("C4", 2, 2, "mailto:test@example.com")];
        rangeLink.TextToDisplay = "Mail";
        rangeLink.ScreenTip = "Send mail";

        return workbook;
    }

    public static void AssertHyperlinks(Workbook workbook)
    {
        var sheet = workbook.Worksheets[0];
        AssertEx.Equal(3, sheet.Hyperlinks.Count);

        var external = sheet.Hyperlinks[0];
        AssertEx.Equal("A1", external.Area);
        AssertEx.Equal("https://example.com/docs?q=1", external.Address);
        AssertEx.Equal("External docs", external.ScreenTip);
        AssertEx.Equal("Docs", external.TextToDisplay);

        var internalLink = sheet.Hyperlinks[1];
        AssertEx.Equal("B2", internalLink.Area);
        AssertEx.Equal("'Target Sheet'!C3", internalLink.Address);
        AssertEx.Equal("Jump to target", internalLink.ScreenTip);
        AssertEx.Equal("Jump", internalLink.TextToDisplay);

        var rangeLink = sheet.Hyperlinks[2];
        AssertEx.Equal("C4:D5", rangeLink.Area);
        AssertEx.Equal("mailto:test@example.com", rangeLink.Address);
        AssertEx.Equal("Send mail", rangeLink.ScreenTip);
        AssertEx.Equal("Mail", rangeLink.TextToDisplay);
    }
}
public static class ZipPackageHelper
{
    public static string ReadEntryText(string packagePath, string entryPath)
    {
        using var archive = ZipFile.OpenRead(packagePath);
        var entry = archive.GetEntry(NormalizeEntryPath(entryPath)) ?? throw new InvalidOperationException($"Missing zip entry '{entryPath}'.");
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    public static bool EntryExists(string packagePath, string entryPath)
    {
        using var archive = ZipFile.OpenRead(packagePath);
        return archive.GetEntry(NormalizeEntryPath(entryPath)) is not null;
    }

    public static void RewriteEntryText(string packagePath, string entryPath, Func<string, string> rewrite)
    {
        using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
        var normalizedPath = NormalizeEntryPath(entryPath);
        var entry = archive.GetEntry(normalizedPath) ?? throw new InvalidOperationException($"Missing zip entry '{entryPath}'.");
        string content;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8, true, 1024, false))
        {
            content = reader.ReadToEnd();
        }

        entry.Delete();
        var replacement = archive.CreateEntry(normalizedPath, CompressionLevel.Optimal);
        using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
        writer.Write(rewrite(content));
    }

    public static void RewriteXmlEntry(string packagePath, string entryPath, Action<XDocument> mutate)
    {
        RewriteEntryText(packagePath, entryPath, delegate(string content)
        {
            var document = XDocument.Parse(content, System.Xml.Linq.LoadOptions.PreserveWhitespace);
            mutate(document);
            return document.Declaration is null ? document.ToString(System.Xml.Linq.SaveOptions.DisableFormatting) : document.Declaration.ToString() + document.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        });
    }

    public static void DeleteEntry(string packagePath, string entryPath)
    {
        using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
        var entry = archive.GetEntry(NormalizeEntryPath(entryPath));
        entry?.Delete();
    }

    public static void MoveEntry(string packagePath, string sourceEntryPath, string destinationEntryPath)
    {
        using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
        var normalizedSourcePath = NormalizeEntryPath(sourceEntryPath);
        var normalizedDestinationPath = NormalizeEntryPath(destinationEntryPath);
        var sourceEntry = archive.GetEntry(normalizedSourcePath) ?? throw new InvalidOperationException($"Missing zip entry '{sourceEntryPath}'.");
        var existingDestination = archive.GetEntry(normalizedDestinationPath);
        existingDestination?.Delete();

        using var buffer = new MemoryStream();
        using (var sourceStream = sourceEntry.Open())
        {
            sourceStream.CopyTo(buffer);
        }

        buffer.Position = 0;
        var destinationEntry = archive.CreateEntry(normalizedDestinationPath, CompressionLevel.Optimal);
        using (var destinationStream = destinationEntry.Open())
        {
            buffer.CopyTo(destinationStream);
        }

        sourceEntry.Delete();
    }
    public static void CreatePackage(string packagePath, IDictionary<string, string> entries)
    {
        using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Create);
        foreach (var pair in entries)
        {
            var entry = archive.CreateEntry(NormalizeEntryPath(pair.Key), CompressionLevel.Optimal);
            using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
            writer.Write(pair.Value);
        }
    }

    private static string NormalizeEntryPath(string entryPath)
    {
        return entryPath.TrimStart('/').Replace('\\', '/');
    }
}



public static class StyleScenarioFactory
{
    private static readonly FillPattern[] FillPatternShowcase =
    {
        FillPattern.Solid,
        FillPattern.MediumGray,
        FillPattern.DarkGray,
        FillPattern.Gray125,
        FillPattern.Gray0625,
        FillPattern.DarkHorizontal,
        FillPattern.DarkVertical,
        FillPattern.DarkDown,
        FillPattern.DarkUp,
        FillPattern.DarkGrid,
        FillPattern.DarkTrellis,
        FillPattern.LightHorizontal,
        FillPattern.LightVertical,
        FillPattern.LightDown,
        FillPattern.LightUp,
        FillPattern.LightGrid,
        FillPattern.LightTrellis,
    };

    private static readonly BorderStyleType[] BorderStyleShowcase =
    {
        BorderStyleType.Thin,
        BorderStyleType.Medium,
        BorderStyleType.Thick,
        BorderStyleType.Dotted,
        BorderStyleType.Dashed,
        BorderStyleType.Double,
        BorderStyleType.Hair,
        BorderStyleType.MediumDashed,
        BorderStyleType.DashDot,
        BorderStyleType.MediumDashDot,
        BorderStyleType.DashDotDot,
        BorderStyleType.MediumDashDotDot,
        BorderStyleType.SlantedDashDot,
    };

    private static readonly HorizontalAlignmentType[] HorizontalAlignmentShowcase =
    {
        HorizontalAlignmentType.General,
        HorizontalAlignmentType.Left,
        HorizontalAlignmentType.Center,
        HorizontalAlignmentType.Right,
        HorizontalAlignmentType.Fill,
        HorizontalAlignmentType.Justify,
        HorizontalAlignmentType.CenterContinuous,
        HorizontalAlignmentType.Distributed,
    };

    private static readonly VerticalAlignmentType[] VerticalAlignmentShowcase =
    {
        VerticalAlignmentType.Bottom,
        VerticalAlignmentType.Center,
        VerticalAlignmentType.Top,
        VerticalAlignmentType.Justify,
        VerticalAlignmentType.Distributed,
    };

    private static readonly string[] NumberFormatShowcase =
    {
        "0",
        "0.00",
        "#,##0.00",
        "0%",
        "0.00E+00",
        "# ?/?",
        "[$-409]#,##0.00",
    };

    private static readonly int[] RotationShowcase =
    {
        0,
        45,
        90,
        135,
        180,
        255,
    };

    private static readonly int[] ReadingOrderShowcase =
    {
        0,
        1,
        2,
    };

    private static readonly double[] FontSizeShowcase =
    {
        8d,
        10d,
        12d,
        14d,
        18d,
        24d,
    };

    private static readonly int[] IndentLevelShowcase =
    {
        0,
        1,
        2,
        3,
        4,
    };

    private const int ShowcaseStartColumn = 5;
    private const int FillPatternRow = 0;
    private const int BorderStyleRow = 1;
    private const int HorizontalAlignmentRow = 2;
    private const int VerticalAlignmentRow = 3;
    private const int FontRow = 4;
    private const int NumberFormatRow = 5;
    private const int BorderSidesRow = 6;
    private const int RotationRow = 7;
    private const int ReadingOrderRow = 8;
    private const int ProtectionRow = 9;
    private const int FontSizeRow = 10;
    private const int IndentRow = 11;
    private const int WrapShrinkRow = 12;

    public static Workbook CreateStyledWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Styled";

        var primaryCell = sheet.Cells["A1"];
        primaryCell.PutValue(1234.567m);
        var primaryStyle = primaryCell.GetStyle();
        ApplyPrimaryStyle(primaryStyle);
        primaryCell.SetStyle(primaryStyle);

        var customCell = sheet.Cells["B2"];
        var customStyle = customCell.GetStyle();
        ApplyCustomNumberStyle(customStyle);
        customCell.SetStyle(customStyle);

        var dateCell = sheet.Cells["C3"];
        dateCell.PutValue(new DateTime(2024, 5, 6, 7, 8, 9));
        var dateStyle = dateCell.GetStyle();
        ApplyDateStyle(dateStyle);
        dateCell.SetStyle(dateStyle);

        var verticalTextCell = sheet.Cells["D4"];
        verticalTextCell.PutValue("Vertical");
        var verticalTextStyle = verticalTextCell.GetStyle();
        ApplyVerticalTextStyle(verticalTextStyle);
        verticalTextCell.SetStyle(verticalTextStyle);

        AddShowcaseLabel(sheet, FillPatternRow, "Fill Patterns");
        AddShowcaseLabel(sheet, BorderStyleRow, "Border Styles");
        AddShowcaseLabel(sheet, HorizontalAlignmentRow, "Horizontal Align");
        AddShowcaseLabel(sheet, VerticalAlignmentRow, "Vertical Align");
        AddShowcaseLabel(sheet, FontRow, "Font Settings");
        AddShowcaseLabel(sheet, NumberFormatRow, "Number Formats");
        AddShowcaseLabel(sheet, BorderSidesRow, "Border Sides");
        AddShowcaseLabel(sheet, RotationRow, "Text Rotation");
        AddShowcaseLabel(sheet, ReadingOrderRow, "Reading Order");
        AddShowcaseLabel(sheet, ProtectionRow, "Protection");
        AddShowcaseLabel(sheet, FontSizeRow, "Font Sizes");
        AddShowcaseLabel(sheet, IndentRow, "Indent Levels");
        AddShowcaseLabel(sheet, WrapShrinkRow, "Wrap And Shrink");

        AddFillPatternShowcase(sheet);
        AddBorderStyleShowcase(sheet);
        AddHorizontalAlignmentShowcase(sheet);
        AddVerticalAlignmentShowcase(sheet);
        AddFontShowcase(sheet);
        AddNumberFormatShowcase(sheet);
        AddBorderSideShowcase(sheet);
        AddRotationShowcase(sheet);
        AddReadingOrderShowcase(sheet);
        AddProtectionShowcase(sheet);
        AddFontSizeShowcase(sheet);
        AddIndentShowcase(sheet);
        AddWrapShrinkShowcase(sheet);

        return workbook;
    }

    public static void ApplyPrimaryStyle(Style style)
    {
        style.Font.Name = "Arial";
        style.Font.Size = 14.5d;
        style.Font.Bold = true;
        style.Font.Italic = true;
        style.Font.Underline = true;
        style.Font.StrikeThrough = true;
        style.Font.Color = Color.FromArgb(255, 17, 34, 51);
        style.Pattern = FillPattern.LightGrid;
        style.ForegroundColor = Color.FromArgb(255, 210, 220, 30);
        style.BackgroundColor = Color.FromArgb(255, 12, 45, 78);
        style.Borders.Left.LineStyle = BorderStyleType.Dotted;
        style.Borders.Left.Color = Color.FromArgb(255, 200, 0, 0);
        style.Borders.Right.LineStyle = BorderStyleType.MediumDashDot;
        style.Borders.Right.Color = Color.FromArgb(255, 240, 120, 0);
        style.Borders.Top.LineStyle = BorderStyleType.Double;
        style.Borders.Top.Color = Color.FromArgb(255, 0, 0, 255);
        style.Borders.Bottom.LineStyle = BorderStyleType.DashDotDot;
        style.Borders.Bottom.Color = Color.FromArgb(255, 0, 120, 0);
        style.Borders.Diagonal.LineStyle = BorderStyleType.SlantedDashDot;
        style.Borders.Diagonal.Color = Color.FromArgb(255, 128, 0, 128);
        style.Borders.DiagonalUp = true;
        style.Borders.DiagonalDown = true;
        style.NumberFormat = "#,##0.00";
        style.HorizontalAlignment = HorizontalAlignmentType.Distributed;
        style.VerticalAlignment = VerticalAlignmentType.Distributed;
        style.WrapText = true;
        style.IndentLevel = 2;
        style.TextRotation = 45;
        style.ShrinkToFit = true;
        style.ReadingOrder = 2;
        style.RelativeIndent = 1;
        style.IsLocked = false;
        style.IsHidden = true;
    }

    public static void ApplyCustomNumberStyle(Style style)
    {
        style.Font.Bold = true;
        style.Borders.Top.LineStyle = BorderStyleType.MediumDashed;
        style.Borders.Top.Color = Color.FromArgb(255, 0, 0, 255);
        style.HorizontalAlignment = HorizontalAlignmentType.CenterContinuous;
        style.VerticalAlignment = VerticalAlignmentType.Top;
        style.NumberFormat = "0.0000";
    }

    public static void ApplyDateStyle(Style style)
    {
        style.Font.Name = "Times New Roman";
        style.Font.Size = 12d;
        style.Font.Color = Color.FromArgb(255, 0, 96, 160);
        style.Pattern = FillPattern.Solid;
        style.ForegroundColor = Color.FromArgb(255, 221, 235, 247);
        style.NumberFormat = "m/d/yyyy h:mm";
        style.HorizontalAlignment = HorizontalAlignmentType.Center;
        style.VerticalAlignment = VerticalAlignmentType.Center;
    }

    public static void ApplyVerticalTextStyle(Style style)
    {
        style.Font.Name = "Consolas";
        style.Font.Size = 10d;
        style.Font.Italic = true;
        style.Pattern = FillPattern.DarkHorizontal;
        style.ForegroundColor = Color.FromArgb(255, 217, 217, 217);
        style.BackgroundColor = Color.FromArgb(255, 255, 255, 255);
        style.Borders.Left.LineStyle = BorderStyleType.Thin;
        style.Borders.Left.Color = Color.FromArgb(255, 64, 64, 64);
        style.HorizontalAlignment = HorizontalAlignmentType.Justify;
        style.VerticalAlignment = VerticalAlignmentType.Justify;
        style.WrapText = true;
        style.TextRotation = 255;
        style.ReadingOrder = 1;
        style.RelativeIndent = 2;
        style.IsLocked = true;
        style.IsHidden = false;
    }

    public static void AssertPrimaryStyle(Style style)
    {
        AssertEx.Equal("Arial", style.Font.Name);
        AssertEx.Equal(14.5d, style.Font.Size);
        AssertEx.True(style.Font.Bold);
        AssertEx.True(style.Font.Italic);
        AssertEx.True(style.Font.Underline);
        AssertEx.True(style.Font.StrikeThrough);
        AssertEx.Equal(Color.FromArgb(255, 17, 34, 51), style.Font.Color);
        AssertEx.Equal(FillPattern.LightGrid, style.Pattern);
        AssertEx.Equal(Color.FromArgb(255, 210, 220, 30), style.ForegroundColor);
        AssertEx.Equal(Color.FromArgb(255, 12, 45, 78), style.BackgroundColor);
        AssertEx.Equal(BorderStyleType.Dotted, style.Borders.Left.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 200, 0, 0), style.Borders.Left.Color);
        AssertEx.Equal(BorderStyleType.MediumDashDot, style.Borders.Right.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 240, 120, 0), style.Borders.Right.Color);
        AssertEx.Equal(BorderStyleType.Double, style.Borders.Top.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 0, 0, 255), style.Borders.Top.Color);
        AssertEx.Equal(BorderStyleType.DashDotDot, style.Borders.Bottom.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 0, 120, 0), style.Borders.Bottom.Color);
        AssertEx.Equal(BorderStyleType.SlantedDashDot, style.Borders.Diagonal.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 128, 0, 128), style.Borders.Diagonal.Color);
        AssertEx.True(style.Borders.DiagonalUp);
        AssertEx.True(style.Borders.DiagonalDown);
        AssertEx.Equal(4, style.Number);
        AssertEx.Null(style.Custom);
        AssertEx.Equal("#,##0.00", style.NumberFormat);
        AssertEx.Equal(HorizontalAlignmentType.Distributed, style.HorizontalAlignment);
        AssertEx.Equal(VerticalAlignmentType.Distributed, style.VerticalAlignment);
        AssertEx.True(style.WrapText);
        AssertEx.Equal(2, style.IndentLevel);
        AssertEx.Equal(45, style.TextRotation);
        AssertEx.True(style.ShrinkToFit);
        AssertEx.Equal(2, style.ReadingOrder);
        AssertEx.Equal(1, style.RelativeIndent);
        AssertEx.False(style.IsLocked);
        AssertEx.True(style.IsHidden);
    }

    public static void AssertCustomNumberStyle(Style style)
    {
        AssertEx.True(style.Font.Bold);
        AssertEx.Equal(BorderStyleType.MediumDashed, style.Borders.Top.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 0, 0, 255), style.Borders.Top.Color);
        AssertEx.Equal(HorizontalAlignmentType.CenterContinuous, style.HorizontalAlignment);
        AssertEx.Equal(VerticalAlignmentType.Top, style.VerticalAlignment);
        AssertEx.Equal("0.0000", style.Custom);
        AssertEx.Equal("0.0000", style.NumberFormat);
    }

    public static void AssertDateStyle(Style style)
    {
        AssertEx.Equal("Times New Roman", style.Font.Name);
        AssertEx.Equal(12d, style.Font.Size);
        AssertEx.Equal(Color.FromArgb(255, 0, 96, 160), style.Font.Color);
        AssertEx.Equal(FillPattern.Solid, style.Pattern);
        AssertEx.Equal(Color.FromArgb(255, 221, 235, 247), style.ForegroundColor);
        AssertEx.Equal("m/d/yyyy h:mm", style.NumberFormat);
        AssertEx.Equal(HorizontalAlignmentType.Center, style.HorizontalAlignment);
        AssertEx.Equal(VerticalAlignmentType.Center, style.VerticalAlignment);
    }

    public static void AssertVerticalTextStyle(Style style)
    {
        AssertEx.Equal("Consolas", style.Font.Name);
        AssertEx.Equal(10d, style.Font.Size);
        AssertEx.True(style.Font.Italic);
        AssertEx.Equal(FillPattern.DarkHorizontal, style.Pattern);
        AssertEx.Equal(Color.FromArgb(255, 217, 217, 217), style.ForegroundColor);
        AssertEx.Equal(Color.FromArgb(255, 255, 255, 255), style.BackgroundColor);
        AssertEx.Equal(BorderStyleType.Thin, style.Borders.Left.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 64, 64, 64), style.Borders.Left.Color);
        AssertEx.Equal(HorizontalAlignmentType.Justify, style.HorizontalAlignment);
        AssertEx.Equal(VerticalAlignmentType.Justify, style.VerticalAlignment);
        AssertEx.True(style.WrapText);
        AssertEx.Equal(255, style.TextRotation);
        AssertEx.Equal(1, style.ReadingOrder);
        AssertEx.Equal(2, style.RelativeIndent);
        AssertEx.True(style.IsLocked);
        AssertEx.False(style.IsHidden);
    }

    public static void AssertStyleShowcase(Workbook workbook)
    {
        var sheet = workbook.Worksheets["Styled"];
        AssertPrimaryStyle(sheet.Cells["A1"].GetStyle());
        AssertCustomNumberStyle(sheet.Cells["B2"].GetStyle());
        AssertDateStyle(sheet.Cells["C3"].GetStyle());
        AssertVerticalTextStyle(sheet.Cells["D4"].GetStyle());

        for (var index = 0; index < FillPatternShowcase.Length; index++)
        {
            var style = sheet.Cells[FillPatternRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal(FillPatternShowcase[index], style.Pattern);
            AssertEx.Equal(GetShowcaseColor(index), style.ForegroundColor);
            AssertEx.Equal(GetShowcaseAccentColor(index), style.BackgroundColor);
        }

        for (var index = 0; index < BorderStyleShowcase.Length; index++)
        {
            var style = sheet.Cells[BorderStyleRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal(BorderStyleShowcase[index], style.Borders.Top.LineStyle);
            AssertEx.Equal(GetShowcaseColor(index), style.Borders.Top.Color);
            AssertEx.Equal(BorderStyleType.Thin, style.Borders.Bottom.LineStyle);
            AssertEx.Equal(GetShowcaseAccentColor(index), style.Borders.Bottom.Color);
        }

        for (var index = 0; index < HorizontalAlignmentShowcase.Length; index++)
        {
            var style = sheet.Cells[HorizontalAlignmentRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal(HorizontalAlignmentShowcase[index], style.HorizontalAlignment);
            AssertEx.True(style.WrapText);
        }

        for (var index = 0; index < VerticalAlignmentShowcase.Length; index++)
        {
            var style = sheet.Cells[VerticalAlignmentRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal(VerticalAlignmentShowcase[index], style.VerticalAlignment);
            AssertEx.True(style.ShrinkToFit);
        }

        AssertFontShowcase(sheet);
        AssertNumberFormatShowcase(sheet);
        AssertBorderSideShowcase(sheet);
        AssertRotationShowcase(sheet);
        AssertReadingOrderShowcase(sheet);
        AssertProtectionShowcase(sheet);
        AssertFontSizeShowcase(sheet);
        AssertIndentShowcase(sheet);
        AssertWrapShrinkShowcase(sheet);
    }

    private static void AddFillPatternShowcase(Worksheet sheet)
    {
        for (var index = 0; index < FillPatternShowcase.Length; index++)
        {
            var cell = sheet.Cells[FillPatternRow, ShowcaseStartColumn + index];
            cell.PutValue("P" + index.ToString());
            var style = cell.GetStyle();
            style.Pattern = FillPatternShowcase[index];
            style.ForegroundColor = GetShowcaseColor(index);
            style.BackgroundColor = GetShowcaseAccentColor(index);
            cell.SetStyle(style);
        }
    }

    private static void AddBorderStyleShowcase(Worksheet sheet)
    {
        for (var index = 0; index < BorderStyleShowcase.Length; index++)
        {
            var cell = sheet.Cells[BorderStyleRow, ShowcaseStartColumn + index];
            cell.PutValue("B" + index.ToString());
            var style = cell.GetStyle();
            style.Borders.Top.LineStyle = BorderStyleShowcase[index];
            style.Borders.Top.Color = GetShowcaseColor(index);
            style.Borders.Bottom.LineStyle = BorderStyleType.Thin;
            style.Borders.Bottom.Color = GetShowcaseAccentColor(index);
            cell.SetStyle(style);
        }
    }

    private static void AddHorizontalAlignmentShowcase(Worksheet sheet)
    {
        for (var index = 0; index < HorizontalAlignmentShowcase.Length; index++)
        {
            var cell = sheet.Cells[HorizontalAlignmentRow, ShowcaseStartColumn + index];
            cell.PutValue("Align " + index.ToString() + " sample");
            var style = cell.GetStyle();
            style.HorizontalAlignment = HorizontalAlignmentShowcase[index];
            style.WrapText = true;
            style.Pattern = FillPattern.Solid;
            style.ForegroundColor = GetShowcaseAccentColor(index);
            cell.SetStyle(style);
        }
    }

    private static void AddVerticalAlignmentShowcase(Worksheet sheet)
    {
        for (var index = 0; index < VerticalAlignmentShowcase.Length; index++)
        {
            var cell = sheet.Cells[VerticalAlignmentRow, ShowcaseStartColumn + index];
            cell.PutValue("Vertical " + index.ToString());
            var style = cell.GetStyle();
            style.VerticalAlignment = VerticalAlignmentShowcase[index];
            style.ShrinkToFit = true;
            style.Pattern = FillPattern.LightTrellis;
            style.ForegroundColor = GetShowcaseColor(index);
            cell.SetStyle(style);
        }
    }

    private static void AddFontShowcase(Worksheet sheet)
    {
        var nameCell = sheet.Cells[FontRow, ShowcaseStartColumn];
        nameCell.PutValue("Cambria 16");
        var nameStyle = nameCell.GetStyle();
        nameStyle.Font.Name = "Cambria";
        nameStyle.Font.Size = 16d;
        nameStyle.Pattern = FillPattern.Solid;
        nameStyle.ForegroundColor = Color.FromArgb(255, 242, 242, 242);
        nameCell.SetStyle(nameStyle);

        var boldCell = sheet.Cells[FontRow, ShowcaseStartColumn + 1];
        boldCell.PutValue("Bold");
        var boldStyle = boldCell.GetStyle();
        boldStyle.Font.Bold = true;
        boldStyle.Font.Size = 12d;
        boldStyle.ForegroundColor = Color.FromArgb(255, 31, 78, 121);
        boldCell.SetStyle(boldStyle);

        var italicCell = sheet.Cells[FontRow, ShowcaseStartColumn + 2];
        italicCell.PutValue("Italic");
        var italicStyle = italicCell.GetStyle();
        italicStyle.Font.Italic = true;
        italicStyle.Font.Size = 12d;
        italicStyle.ForegroundColor = Color.FromArgb(255, 128, 100, 162);
        italicCell.SetStyle(italicStyle);

        var underlineCell = sheet.Cells[FontRow, ShowcaseStartColumn + 3];
        underlineCell.PutValue("Underline");
        var underlineStyle = underlineCell.GetStyle();
        underlineStyle.Font.Underline = true;
        underlineStyle.Font.Color = Color.FromArgb(255, 0, 112, 192);
        underlineCell.SetStyle(underlineStyle);

        var strikeCell = sheet.Cells[FontRow, ShowcaseStartColumn + 4];
        strikeCell.PutValue("Strike");
        var strikeStyle = strikeCell.GetStyle();
        strikeStyle.Font.StrikeThrough = true;
        strikeStyle.Font.Color = Color.FromArgb(255, 192, 0, 0);
        strikeCell.SetStyle(strikeStyle);

        var colorCell = sheet.Cells[FontRow, ShowcaseStartColumn + 5];
        colorCell.PutValue("Color");
        var colorStyle = colorCell.GetStyle();
        colorStyle.Font.Bold = true;
        colorStyle.Font.Color = Color.FromArgb(255, 0, 176, 80);
        colorStyle.Pattern = FillPattern.Solid;
        colorStyle.ForegroundColor = Color.FromArgb(255, 226, 239, 218);
        colorCell.SetStyle(colorStyle);
    }

    private static void AddNumberFormatShowcase(Worksheet sheet)
    {
        for (var index = 0; index < NumberFormatShowcase.Length; index++)
        {
            var cell = sheet.Cells[NumberFormatRow, ShowcaseStartColumn + index];
            if (index == 3)
            {
                cell.PutValue(0.375d);
            }
            else if (index == 5)
            {
                cell.PutValue(12.75d);
            }
            else
            {
                cell.PutValue(1234.567d + index);
            }

            var style = cell.GetStyle();
            style.NumberFormat = NumberFormatShowcase[index];
            style.HorizontalAlignment = HorizontalAlignmentType.Right;
            style.Pattern = FillPattern.Solid;
            style.ForegroundColor = GetShowcaseAccentColor(index + 10);
            cell.SetStyle(style);
        }
    }

    private static void AddBorderSideShowcase(Worksheet sheet)
    {
        var leftCell = sheet.Cells[BorderSidesRow, ShowcaseStartColumn];
        leftCell.PutValue("Left");
        var leftStyle = leftCell.GetStyle();
        leftStyle.Borders.Left.LineStyle = BorderStyleType.Thick;
        leftStyle.Borders.Left.Color = Color.FromArgb(255, 192, 0, 0);
        leftStyle.Pattern = FillPattern.Solid;
        leftStyle.ForegroundColor = Color.FromArgb(255, 255, 242, 204);
        leftCell.SetStyle(leftStyle);

        var rightCell = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 1];
        rightCell.PutValue("Right");
        var rightStyle = rightCell.GetStyle();
        rightStyle.Borders.Right.LineStyle = BorderStyleType.Double;
        rightStyle.Borders.Right.Color = Color.FromArgb(255, 0, 112, 192);
        rightStyle.Pattern = FillPattern.Solid;
        rightStyle.ForegroundColor = Color.FromArgb(255, 221, 235, 247);
        rightCell.SetStyle(rightStyle);

        var topCell = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 2];
        topCell.PutValue("Top");
        var topStyle = topCell.GetStyle();
        topStyle.Borders.Top.LineStyle = BorderStyleType.DashDot;
        topStyle.Borders.Top.Color = Color.FromArgb(255, 112, 48, 160);
        topStyle.Pattern = FillPattern.Solid;
        topStyle.ForegroundColor = Color.FromArgb(255, 234, 209, 220);
        topCell.SetStyle(topStyle);

        var bottomCell = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 3];
        bottomCell.PutValue("Bottom");
        var bottomStyle = bottomCell.GetStyle();
        bottomStyle.Borders.Bottom.LineStyle = BorderStyleType.MediumDashed;
        bottomStyle.Borders.Bottom.Color = Color.FromArgb(255, 0, 176, 80);
        bottomStyle.Pattern = FillPattern.Solid;
        bottomStyle.ForegroundColor = Color.FromArgb(255, 226, 239, 218);
        bottomCell.SetStyle(bottomStyle);

        var diagonalUpCell = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 4];
        diagonalUpCell.PutValue("Diag Up");
        var diagonalUpStyle = diagonalUpCell.GetStyle();
        diagonalUpStyle.Borders.Diagonal.LineStyle = BorderStyleType.SlantedDashDot;
        diagonalUpStyle.Borders.Diagonal.Color = Color.FromArgb(255, 255, 0, 0);
        diagonalUpStyle.Borders.DiagonalUp = true;
        diagonalUpStyle.Pattern = FillPattern.Solid;
        diagonalUpStyle.ForegroundColor = Color.FromArgb(255, 252, 228, 214);
        diagonalUpCell.SetStyle(diagonalUpStyle);

        var diagonalDownCell = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 5];
        diagonalDownCell.PutValue("Diag Down");
        var diagonalDownStyle = diagonalDownCell.GetStyle();
        diagonalDownStyle.Borders.Diagonal.LineStyle = BorderStyleType.SlantedDashDot;
        diagonalDownStyle.Borders.Diagonal.Color = Color.FromArgb(255, 0, 176, 240);
        diagonalDownStyle.Borders.DiagonalDown = true;
        diagonalDownStyle.Pattern = FillPattern.Solid;
        diagonalDownStyle.ForegroundColor = Color.FromArgb(255, 222, 234, 246);
        diagonalDownCell.SetStyle(diagonalDownStyle);
    }

    private static void AddRotationShowcase(Worksheet sheet)
    {
        for (var index = 0; index < RotationShowcase.Length; index++)
        {
            var cell = sheet.Cells[RotationRow, ShowcaseStartColumn + index];
            cell.PutValue("Rot " + RotationShowcase[index].ToString());
            var style = cell.GetStyle();
            style.TextRotation = RotationShowcase[index];
            style.HorizontalAlignment = HorizontalAlignmentType.Center;
            style.VerticalAlignment = VerticalAlignmentType.Center;
            style.WrapText = true;
            style.Pattern = FillPattern.Solid;
            style.ForegroundColor = GetShowcaseColor(index + 20);
            cell.SetStyle(style);
        }
    }

    private static void AddReadingOrderShowcase(Worksheet sheet)
    {
        for (var index = 0; index < ReadingOrderShowcase.Length; index++)
        {
            var cell = sheet.Cells[ReadingOrderRow, ShowcaseStartColumn + index];
            cell.PutValue("Order " + ReadingOrderShowcase[index].ToString());
            var style = cell.GetStyle();
            style.ReadingOrder = ReadingOrderShowcase[index];
            style.IndentLevel = index;
            if (ReadingOrderShowcase[index] > 0)
            {
                style.RelativeIndent = index + 1;
            }
            style.HorizontalAlignment = HorizontalAlignmentType.Distributed;
            style.Pattern = FillPattern.Solid;
            style.ForegroundColor = GetShowcaseAccentColor(index + 20);
            cell.SetStyle(style);
        }
    }

    private static void AddProtectionShowcase(Worksheet sheet)
    {
        var lockedCell = sheet.Cells[ProtectionRow, ShowcaseStartColumn];
        lockedCell.PutValue("Locked");
        var lockedStyle = lockedCell.GetStyle();
        lockedStyle.IsLocked = true;
        lockedStyle.IsHidden = false;
        lockedStyle.Pattern = FillPattern.Solid;
        lockedStyle.ForegroundColor = Color.FromArgb(255, 226, 239, 218);
        lockedCell.SetStyle(lockedStyle);

        var unlockedCell = sheet.Cells[ProtectionRow, ShowcaseStartColumn + 1];
        unlockedCell.PutValue("Unlocked");
        var unlockedStyle = unlockedCell.GetStyle();
        unlockedStyle.IsLocked = false;
        unlockedStyle.IsHidden = false;
        unlockedStyle.Pattern = FillPattern.Solid;
        unlockedStyle.ForegroundColor = Color.FromArgb(255, 255, 242, 204);
        unlockedCell.SetStyle(unlockedStyle);

        var hiddenCell = sheet.Cells[ProtectionRow, ShowcaseStartColumn + 2];
        hiddenCell.PutValue("Hidden");
        var hiddenStyle = hiddenCell.GetStyle();
        hiddenStyle.IsLocked = true;
        hiddenStyle.IsHidden = true;
        hiddenStyle.Pattern = FillPattern.Solid;
        hiddenStyle.ForegroundColor = Color.FromArgb(255, 242, 220, 219);
        hiddenCell.SetStyle(hiddenStyle);
    }

    private static void AddFontSizeShowcase(Worksheet sheet)
    {
        for (var index = 0; index < FontSizeShowcase.Length; index++)
        {
            var cell = sheet.Cells[FontSizeRow, ShowcaseStartColumn + index];
            cell.PutValue("Size " + FontSizeShowcase[index].ToString("0"));
            var style = cell.GetStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = FontSizeShowcase[index];
            style.Pattern = FillPattern.Solid;
            style.ForegroundColor = GetShowcaseAccentColor(index + 30);
            cell.SetStyle(style);
        }
    }

    private static void AddIndentShowcase(Worksheet sheet)
    {
        for (var index = 0; index < IndentLevelShowcase.Length; index++)
        {
            var cell = sheet.Cells[IndentRow, ShowcaseStartColumn + index];
            cell.PutValue("Indent " + IndentLevelShowcase[index].ToString());
            var style = cell.GetStyle();
            style.HorizontalAlignment = HorizontalAlignmentType.Left;
            style.IndentLevel = IndentLevelShowcase[index];
            style.RelativeIndent = IndentLevelShowcase[index];
            style.Pattern = FillPattern.Solid;
            style.ForegroundColor = GetShowcaseColor(index + 30);
            cell.SetStyle(style);
        }
    }

    private static void AddWrapShrinkShowcase(Worksheet sheet)
    {
        var wrapCell = sheet.Cells[WrapShrinkRow, ShowcaseStartColumn];
        wrapCell.PutValue("Wrap text sample");
        var wrapStyle = wrapCell.GetStyle();
        wrapStyle.WrapText = true;
        wrapStyle.HorizontalAlignment = HorizontalAlignmentType.Justify;
        wrapStyle.Pattern = FillPattern.Solid;
        wrapStyle.ForegroundColor = Color.FromArgb(255, 252, 228, 214);
        wrapCell.SetStyle(wrapStyle);

        var shrinkCell = sheet.Cells[WrapShrinkRow, ShowcaseStartColumn + 1];
        shrinkCell.PutValue("Shrink to fit sample");
        var shrinkStyle = shrinkCell.GetStyle();
        shrinkStyle.ShrinkToFit = true;
        shrinkStyle.HorizontalAlignment = HorizontalAlignmentType.Center;
        shrinkStyle.Pattern = FillPattern.Solid;
        shrinkStyle.ForegroundColor = Color.FromArgb(255, 221, 235, 247);
        shrinkCell.SetStyle(shrinkStyle);

        var wrapShrinkCell = sheet.Cells[WrapShrinkRow, ShowcaseStartColumn + 2];
        wrapShrinkCell.PutValue("Wrap and shrink");
        var wrapShrinkStyle = wrapShrinkCell.GetStyle();
        wrapShrinkStyle.WrapText = true;
        wrapShrinkStyle.ShrinkToFit = true;
        wrapShrinkStyle.HorizontalAlignment = HorizontalAlignmentType.Distributed;
        wrapShrinkStyle.VerticalAlignment = VerticalAlignmentType.Center;
        wrapShrinkStyle.Pattern = FillPattern.Solid;
        wrapShrinkStyle.ForegroundColor = Color.FromArgb(255, 226, 239, 218);
        wrapShrinkCell.SetStyle(wrapShrinkStyle);

        var justifyCell = sheet.Cells[WrapShrinkRow, ShowcaseStartColumn + 3];
        justifyCell.PutValue("Distributed indent");
        var justifyStyle = justifyCell.GetStyle();
        justifyStyle.WrapText = true;
        justifyStyle.HorizontalAlignment = HorizontalAlignmentType.Distributed;
        justifyStyle.IndentLevel = 3;
        justifyStyle.RelativeIndent = 2;
        justifyStyle.Pattern = FillPattern.Solid;
        justifyStyle.ForegroundColor = Color.FromArgb(255, 242, 242, 242);
        justifyCell.SetStyle(justifyStyle);
    }

    private static void AssertFontShowcase(Worksheet sheet)
    {
        var nameStyle = sheet.Cells[FontRow, ShowcaseStartColumn].GetStyle();
        AssertEx.Equal("Cambria", nameStyle.Font.Name);
        AssertEx.Equal(16d, nameStyle.Font.Size);
        AssertEx.Equal(FillPattern.Solid, nameStyle.Pattern);

        var boldStyle = sheet.Cells[FontRow, ShowcaseStartColumn + 1].GetStyle();
        AssertEx.True(boldStyle.Font.Bold);
        AssertEx.Equal(12d, boldStyle.Font.Size);
        AssertEx.Equal(Color.FromArgb(255, 31, 78, 121), boldStyle.ForegroundColor);

        var italicStyle = sheet.Cells[FontRow, ShowcaseStartColumn + 2].GetStyle();
        AssertEx.True(italicStyle.Font.Italic);
        AssertEx.Equal(12d, italicStyle.Font.Size);
        AssertEx.Equal(Color.FromArgb(255, 128, 100, 162), italicStyle.ForegroundColor);

        var underlineStyle = sheet.Cells[FontRow, ShowcaseStartColumn + 3].GetStyle();
        AssertEx.True(underlineStyle.Font.Underline);
        AssertEx.Equal(Color.FromArgb(255, 0, 112, 192), underlineStyle.Font.Color);

        var strikeStyle = sheet.Cells[FontRow, ShowcaseStartColumn + 4].GetStyle();
        AssertEx.True(strikeStyle.Font.StrikeThrough);
        AssertEx.Equal(Color.FromArgb(255, 192, 0, 0), strikeStyle.Font.Color);

        var colorStyle = sheet.Cells[FontRow, ShowcaseStartColumn + 5].GetStyle();
        AssertEx.True(colorStyle.Font.Bold);
        AssertEx.Equal(Color.FromArgb(255, 0, 176, 80), colorStyle.Font.Color);
        AssertEx.Equal(FillPattern.Solid, colorStyle.Pattern);
        AssertEx.Equal(Color.FromArgb(255, 226, 239, 218), colorStyle.ForegroundColor);
    }

    private static void AssertNumberFormatShowcase(Worksheet sheet)
    {
        for (var index = 0; index < NumberFormatShowcase.Length; index++)
        {
            var style = sheet.Cells[NumberFormatRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal(NumberFormatShowcase[index], style.NumberFormat);
            AssertEx.Equal(HorizontalAlignmentType.Right, style.HorizontalAlignment);
            AssertEx.Equal(FillPattern.Solid, style.Pattern);
            AssertEx.Equal(GetShowcaseAccentColor(index + 10), style.ForegroundColor);
        }
    }

    private static void AssertBorderSideShowcase(Worksheet sheet)
    {
        var leftStyle = sheet.Cells[BorderSidesRow, ShowcaseStartColumn].GetStyle();
        AssertEx.Equal(BorderStyleType.Thick, leftStyle.Borders.Left.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 192, 0, 0), leftStyle.Borders.Left.Color);

        var rightStyle = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 1].GetStyle();
        AssertEx.Equal(BorderStyleType.Double, rightStyle.Borders.Right.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 0, 112, 192), rightStyle.Borders.Right.Color);

        var topStyle = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 2].GetStyle();
        AssertEx.Equal(BorderStyleType.DashDot, topStyle.Borders.Top.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 112, 48, 160), topStyle.Borders.Top.Color);

        var bottomStyle = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 3].GetStyle();
        AssertEx.Equal(BorderStyleType.MediumDashed, bottomStyle.Borders.Bottom.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 0, 176, 80), bottomStyle.Borders.Bottom.Color);

        var diagonalUpStyle = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 4].GetStyle();
        AssertEx.Equal(BorderStyleType.SlantedDashDot, diagonalUpStyle.Borders.Diagonal.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 255, 0, 0), diagonalUpStyle.Borders.Diagonal.Color);
        AssertEx.True(diagonalUpStyle.Borders.DiagonalUp);
        AssertEx.False(diagonalUpStyle.Borders.DiagonalDown);

        var diagonalDownStyle = sheet.Cells[BorderSidesRow, ShowcaseStartColumn + 5].GetStyle();
        AssertEx.Equal(BorderStyleType.SlantedDashDot, diagonalDownStyle.Borders.Diagonal.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 0, 176, 240), diagonalDownStyle.Borders.Diagonal.Color);
        AssertEx.False(diagonalDownStyle.Borders.DiagonalUp);
        AssertEx.True(diagonalDownStyle.Borders.DiagonalDown);
    }

    private static void AssertRotationShowcase(Worksheet sheet)
    {
        for (var index = 0; index < RotationShowcase.Length; index++)
        {
            var style = sheet.Cells[RotationRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal(RotationShowcase[index], style.TextRotation);
            AssertEx.Equal(HorizontalAlignmentType.Center, style.HorizontalAlignment);
            AssertEx.Equal(VerticalAlignmentType.Center, style.VerticalAlignment);
            AssertEx.True(style.WrapText);
        }
    }

    private static void AssertReadingOrderShowcase(Worksheet sheet)
    {
        for (var index = 0; index < ReadingOrderShowcase.Length; index++)
        {
            var style = sheet.Cells[ReadingOrderRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal(ReadingOrderShowcase[index], style.ReadingOrder);
            AssertEx.Equal(index, style.IndentLevel);
            AssertEx.Equal(ReadingOrderShowcase[index] > 0 ? index + 1 : 0, style.RelativeIndent);
            AssertEx.Equal(HorizontalAlignmentType.Distributed, style.HorizontalAlignment);
            AssertEx.Equal(FillPattern.Solid, style.Pattern);
            AssertEx.Equal(GetShowcaseAccentColor(index + 20), style.ForegroundColor);
        }
    }

    private static void AssertProtectionShowcase(Worksheet sheet)
    {
        var lockedStyle = sheet.Cells[ProtectionRow, ShowcaseStartColumn].GetStyle();
        AssertEx.True(lockedStyle.IsLocked);
        AssertEx.False(lockedStyle.IsHidden);

        var unlockedStyle = sheet.Cells[ProtectionRow, ShowcaseStartColumn + 1].GetStyle();
        AssertEx.False(unlockedStyle.IsLocked);
        AssertEx.False(unlockedStyle.IsHidden);

        var hiddenStyle = sheet.Cells[ProtectionRow, ShowcaseStartColumn + 2].GetStyle();
        AssertEx.True(hiddenStyle.IsLocked);
        AssertEx.True(hiddenStyle.IsHidden);
    }

    private static void AssertFontSizeShowcase(Worksheet sheet)
    {
        for (var index = 0; index < FontSizeShowcase.Length; index++)
        {
            var style = sheet.Cells[FontSizeRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal("Calibri", style.Font.Name);
            AssertEx.Equal(FontSizeShowcase[index], style.Font.Size);
            AssertEx.Equal(FillPattern.Solid, style.Pattern);
            AssertEx.Equal(GetShowcaseAccentColor(index + 30), style.ForegroundColor);
        }
    }

    private static void AssertIndentShowcase(Worksheet sheet)
    {
        for (var index = 0; index < IndentLevelShowcase.Length; index++)
        {
            var style = sheet.Cells[IndentRow, ShowcaseStartColumn + index].GetStyle();
            AssertEx.Equal(HorizontalAlignmentType.Left, style.HorizontalAlignment);
            AssertEx.Equal(IndentLevelShowcase[index], style.IndentLevel);
            AssertEx.Equal(FillPattern.Solid, style.Pattern);
            AssertEx.Equal(GetShowcaseColor(index + 30), style.ForegroundColor);
        }
    }

    private static void AssertWrapShrinkShowcase(Worksheet sheet)
    {
        var wrapStyle = sheet.Cells[WrapShrinkRow, ShowcaseStartColumn].GetStyle();
        AssertEx.True(wrapStyle.WrapText);
        AssertEx.Equal(HorizontalAlignmentType.Justify, wrapStyle.HorizontalAlignment);

        var shrinkStyle = sheet.Cells[WrapShrinkRow, ShowcaseStartColumn + 1].GetStyle();
        AssertEx.True(shrinkStyle.ShrinkToFit);
        AssertEx.Equal(HorizontalAlignmentType.Center, shrinkStyle.HorizontalAlignment);

        var wrapShrinkStyle = sheet.Cells[WrapShrinkRow, ShowcaseStartColumn + 2].GetStyle();
        AssertEx.True(wrapShrinkStyle.WrapText);
        AssertEx.True(wrapShrinkStyle.ShrinkToFit);
        AssertEx.Equal(HorizontalAlignmentType.Distributed, wrapShrinkStyle.HorizontalAlignment);
        AssertEx.Equal(VerticalAlignmentType.Center, wrapShrinkStyle.VerticalAlignment);

        var justifyStyle = sheet.Cells[WrapShrinkRow, ShowcaseStartColumn + 3].GetStyle();
        AssertEx.True(justifyStyle.WrapText);
        AssertEx.Equal(HorizontalAlignmentType.Distributed, justifyStyle.HorizontalAlignment);
        AssertEx.Equal(3, justifyStyle.IndentLevel);
        AssertEx.Equal(0, justifyStyle.RelativeIndent);
    }

    private static void AddShowcaseLabel(Worksheet sheet, int rowIndex, string label)
    {
        var cell = sheet.Cells[rowIndex, ShowcaseStartColumn - 1];
        cell.PutValue(label);
        var style = cell.GetStyle();
        style.Font.Bold = true;
        style.Pattern = FillPattern.Solid;
        style.ForegroundColor = Color.FromArgb(255, 217, 225, 242);
        style.Borders.Bottom.LineStyle = BorderStyleType.Thin;
        style.Borders.Bottom.Color = Color.FromArgb(255, 79, 129, 189);
        cell.SetStyle(style);
    }

    private static Color GetShowcaseColor(int index)
    {
        return Color.FromArgb(255,
            (40 + (index * 29)) % 256,
            (90 + (index * 37)) % 256,
            (140 + (index * 43)) % 256);
    }

    private static Color GetShowcaseAccentColor(int index)
    {
        return Color.FromArgb(255,
            (220 + (index * 17)) % 256,
            (180 + (index * 19)) % 256,
            (120 + (index * 23)) % 256);
    }
}
public static class ValidationScenarioFactory
{
    public static Workbook CreateValidationWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Validation Sheet";

        sheet.Cells["A1"].PutValue("Open");
        sheet.Cells["B2"].PutValue(5);
        sheet.Cells["C3"].PutValue(7);
        sheet.Cells["E2"].PutValue(8);
        sheet.Cells["G1"].PutValue("ABCDE");

        var listIndex = sheet.Validations.Add(CellArea.CreateCellArea("A1", "A3"));
        var listValidation = sheet.Validations[listIndex];
        listValidation.Type = ValidationType.List;
        listValidation.Formula1 = "\"Open,Closed\"";
        listValidation.IgnoreBlank = true;
        listValidation.InCellDropDown = true;
        listValidation.ShowInput = true;
        listValidation.InputTitle = "Status";
        listValidation.InputMessage = "Pick a status";
        listValidation.ShowError = true;
        listValidation.ErrorTitle = "Invalid";
        listValidation.ErrorMessage = "Choose from the list";

        var decimalIndex = sheet.Validations.Add(CellArea.CreateCellArea("B2", "C3"));
        var decimalValidation = sheet.Validations[decimalIndex];
        decimalValidation.Type = ValidationType.Decimal;
        decimalValidation.Operator = OperatorType.Between;
        decimalValidation.Formula1 = "1.5";
        decimalValidation.Formula2 = "9.5";
        decimalValidation.IgnoreBlank = false;
        decimalValidation.InCellDropDown = false;
        decimalValidation.ShowError = true;
        decimalValidation.ErrorTitle = "Range";
        decimalValidation.ErrorMessage = "Enter 1.5-9.5";
        decimalValidation.AddArea(CellArea.CreateCellArea("E2", "E3"));

        var customIndex = sheet.Validations.Add(CellArea.CreateCellArea("G1", "G1"));
        var customValidation = sheet.Validations[customIndex];
        customValidation.Type = ValidationType.Custom;
        customValidation.AlertStyle = ValidationAlertType.Warning;
        customValidation.Formula1 = "LEN(G1)<=5";
        customValidation.ShowInput = true;
        customValidation.InputTitle = "Code";
        customValidation.InputMessage = "Up to 5 chars";

        return workbook;
    }

    public static void AssertValidations(Workbook workbook)
    {
        var sheet = workbook.Worksheets["Validation Sheet"];
        AssertEx.Equal(3, sheet.Validations.Count);

        var listValidation = GetValidationByArea(sheet.Validations, "A1:A3");
        AssertEx.Equal(ValidationType.List, listValidation.Type);
        AssertEx.Equal(OperatorType.None, listValidation.Operator);
        AssertEx.Equal(1, listValidation.Areas.Count);
        AssertArea(listValidation.Areas[0], 0, 0, 3, 1);
        AssertEx.Equal("\"Open,Closed\"", listValidation.Formula1);
        AssertEx.Equal(string.Empty, listValidation.Formula2);
        AssertEx.True(listValidation.IgnoreBlank);
        AssertEx.True(listValidation.InCellDropDown);
        AssertEx.True(listValidation.ShowInput);
        AssertEx.True(listValidation.ShowError);
        AssertEx.Equal("Status", listValidation.InputTitle);
        AssertEx.Equal("Pick a status", listValidation.InputMessage);
        AssertEx.Equal("Invalid", listValidation.ErrorTitle);
        AssertEx.Equal("Choose from the list", listValidation.ErrorMessage);

        var decimalValidation = GetValidationByArea(sheet.Validations, "B2:C3 E2:E3");
        AssertEx.Equal(ValidationType.Decimal, decimalValidation.Type);
        AssertEx.Equal(OperatorType.Between, decimalValidation.Operator);
        AssertEx.Equal(2, decimalValidation.Areas.Count);
        AssertArea(decimalValidation.Areas[0], 1, 1, 2, 2);
        AssertArea(decimalValidation.Areas[1], 1, 4, 2, 1);
        AssertEx.Equal("1.5", decimalValidation.Formula1);
        AssertEx.Equal("9.5", decimalValidation.Formula2);
        AssertEx.False(decimalValidation.IgnoreBlank);
        AssertEx.False(decimalValidation.InCellDropDown);
        AssertEx.False(decimalValidation.ShowInput);
        AssertEx.True(decimalValidation.ShowError);
        AssertEx.Equal("Range", decimalValidation.ErrorTitle);
        AssertEx.Equal("Enter 1.5-9.5", decimalValidation.ErrorMessage);

        var customValidation = GetValidationByArea(sheet.Validations, "G1");
        AssertEx.Equal(ValidationType.Custom, customValidation.Type);
        AssertEx.Equal(ValidationAlertType.Warning, customValidation.AlertStyle);
        AssertEx.Equal(OperatorType.None, customValidation.Operator);
        AssertEx.Equal(1, customValidation.Areas.Count);
        AssertArea(customValidation.Areas[0], 0, 6, 1, 1);
        AssertEx.Equal("LEN(G1)<=5", customValidation.Formula1);
        AssertEx.Equal(string.Empty, customValidation.Formula2);
        AssertEx.True(customValidation.ShowInput);
        AssertEx.False(customValidation.ShowError);
        AssertEx.Equal("Code", customValidation.InputTitle);
        AssertEx.Equal("Up to 5 chars", customValidation.InputMessage);

        var a1Validation = sheet.Validations.GetValidationInCell(0, 0);
        AssertEx.NotNull(a1Validation);
        AssertEx.Equal(ValidationType.List, a1Validation!.Type);

        var c2Validation = sheet.Validations.GetValidationInCell(1, 2);
        AssertEx.NotNull(c2Validation);
        AssertEx.Equal(ValidationType.Decimal, c2Validation!.Type);

        var missingValidation = sheet.Validations.GetValidationInCell(9, 9);
        AssertEx.Null(missingValidation);
    }

    private static Validation GetValidationByArea(ValidationCollection validations, string expectedArea)
    {
        for (var index = 0; index < validations.Count; index++)
        {
            var validation = validations[index];
            if (string.Equals(BuildAreaKey(validation), expectedArea, StringComparison.Ordinal))
            {
                return validation;
            }
        }

        throw new InvalidOperationException("Validation with area '" + expectedArea + "' was not found.");
    }

    private static string BuildAreaKey(Validation validation)
    {
        var areas = new List<string>(validation.Areas.Count);
        for (var index = 0; index < validation.Areas.Count; index++)
        {
            var area = validation.Areas[index];
            var startCell = BuildCellName(area.FirstRow, area.FirstColumn);
            if (area.TotalRows == 1 && area.TotalColumns == 1)
            {
                areas.Add(startCell);
                continue;
            }

            var endCell = BuildCellName(area.FirstRow + area.TotalRows - 1, area.FirstColumn + area.TotalColumns - 1);
            areas.Add(startCell + ":" + endCell);
        }

        return string.Join(" ", areas);
    }

    private static string BuildCellName(int rowIndex, int columnIndex)
    {
        var dividend = columnIndex + 1;
        var name = string.Empty;
        while (dividend > 0)
        {
            var remainder = (dividend - 1) % 26;
            name = (char)('A' + remainder) + name;
            dividend = (dividend - remainder - 1) / 26;
        }

        return name + (rowIndex + 1).ToString();
    }

    private static void AssertArea(CellArea area, int firstRow, int firstColumn, int totalRows, int totalColumns)
    {
        AssertEx.Equal(firstRow, area.FirstRow);
        AssertEx.Equal(firstColumn, area.FirstColumn);
        AssertEx.Equal(totalRows, area.TotalRows);
        AssertEx.Equal(totalColumns, area.TotalColumns);
    }
}


public static class ConditionalFormattingScenarioFactory
{
    public static Workbook CreateConditionalFormattingWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Conditional Formatting";

        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(5);
        sheet.Cells["A3"].PutValue(10);
        sheet.Cells["A4"].PutValue(15);
        sheet.Cells["A5"].PutValue(20);
        sheet.Cells["C1"].PutValue(11);
        sheet.Cells["C2"].PutValue(8);
        sheet.Cells["C3"].PutValue(15);

        var primaryIndex = sheet.ConditionalFormattings.Add();
        var primary = sheet.ConditionalFormattings[primaryIndex];
        primary.AddArea(CellArea.CreateCellArea("A1", "A5"));
        var betweenIndex = primary.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "5", "15");
        var between = primary[betweenIndex];
        between.StopIfTrue = true;
        between.Priority = 1;
        var betweenStyle = between.Style;
        betweenStyle.Pattern = FillPattern.Solid;
        betweenStyle.ForegroundColor = Color.FromArgb(255, 255, 199, 206);
        betweenStyle.Font.Bold = true;
        betweenStyle.Font.Color = Color.FromArgb(255, 156, 0, 6);
        between.Style = betweenStyle;

        var expressionIndex = primary.AddCondition(FormatConditionType.Expression, OperatorType.None, "MOD(A1,2)=0", string.Empty);
        var expression = primary[expressionIndex];
        expression.Priority = 2;
        var expressionStyle = expression.Style;
        expressionStyle.Font.Italic = true;
        expressionStyle.Font.Color = Color.FromArgb(255, 0, 0, 255);
        expressionStyle.Borders.Bottom.LineStyle = BorderStyleType.Thin;
        expressionStyle.Borders.Bottom.Color = Color.FromArgb(255, 0, 0, 255);
        expression.Style = expressionStyle;

        var secondaryIndex = sheet.ConditionalFormattings.Add();
        var secondary = sheet.ConditionalFormattings[secondaryIndex];
        secondary.AddArea(CellArea.CreateCellArea("C1", "C3"));
        var greaterThanIndex = secondary.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", string.Empty);
        var greaterThan = secondary[greaterThanIndex];
        greaterThan.Priority = 3;
        var greaterThanStyle = greaterThan.Style;
        greaterThanStyle.Pattern = FillPattern.Solid;
        greaterThanStyle.ForegroundColor = Color.FromArgb(255, 198, 239, 206);
        greaterThanStyle.Font.Color = Color.FromArgb(255, 0, 97, 0);
        greaterThan.Style = greaterThanStyle;

        return workbook;
    }

    public static void AssertConditionalFormattings(Workbook workbook)
    {
        var sheet = workbook.Worksheets["Conditional Formatting"];
        AssertEx.Equal(2, sheet.ConditionalFormattings.Count);

        var primary = GetCollectionByArea(sheet.ConditionalFormattings, "A1:A5");
        AssertEx.Equal(2, primary.Count);
        AssertEx.Equal(1, primary.RangeCount);
        AssertArea(primary.GetCellArea(0), 0, 0, 5, 1);

        var between = primary[0];
        AssertEx.Equal(FormatConditionType.CellValue, between.Type);
        AssertEx.Equal(OperatorType.Between, between.Operator);
        AssertEx.Equal("5", between.Formula1);
        AssertEx.Equal("15", between.Formula2);
        AssertEx.Equal(1, between.Priority);
        AssertEx.True(between.StopIfTrue);
        AssertEx.Equal(FillPattern.Solid, between.Style.Pattern);
        AssertEx.Equal(Color.FromArgb(255, 255, 199, 206), between.Style.ForegroundColor);
        AssertEx.True(between.Style.Font.Bold);
        AssertEx.Equal(Color.FromArgb(255, 156, 0, 6), between.Style.Font.Color);

        var expression = primary[1];
        AssertEx.Equal(FormatConditionType.Expression, expression.Type);
        AssertEx.Equal(OperatorType.None, expression.Operator);
        AssertEx.Equal("MOD(A1,2)=0", expression.Formula1);
        AssertEx.Equal(string.Empty, expression.Formula2);
        AssertEx.Equal(2, expression.Priority);
        AssertEx.False(expression.StopIfTrue);
        AssertEx.True(expression.Style.Font.Italic);
        AssertEx.Equal(Color.FromArgb(255, 0, 0, 255), expression.Style.Font.Color);
        AssertEx.Equal(BorderStyleType.Thin, expression.Style.Borders.Bottom.LineStyle);
        AssertEx.Equal(Color.FromArgb(255, 0, 0, 255), expression.Style.Borders.Bottom.Color);

        var secondary = GetCollectionByArea(sheet.ConditionalFormattings, "C1:C3");
        AssertEx.Equal(1, secondary.Count);
        AssertEx.Equal(1, secondary.RangeCount);
        AssertArea(secondary.GetCellArea(0), 0, 2, 3, 1);
        var greaterThan = secondary[0];
        AssertEx.Equal(FormatConditionType.CellValue, greaterThan.Type);
        AssertEx.Equal(OperatorType.GreaterThan, greaterThan.Operator);
        AssertEx.Equal("10", greaterThan.Formula1);
        AssertEx.Equal(string.Empty, greaterThan.Formula2);
        AssertEx.Equal(3, greaterThan.Priority);
        AssertEx.Equal(FillPattern.Solid, greaterThan.Style.Pattern);
        AssertEx.Equal(Color.FromArgb(255, 198, 239, 206), greaterThan.Style.ForegroundColor);
        AssertEx.Equal(Color.FromArgb(255, 0, 97, 0), greaterThan.Style.Font.Color);
    }

    private static FormatConditionCollection GetCollectionByArea(ConditionalFormattingCollection collections, string expectedArea)
    {
        for (var index = 0; index < collections.Count; index++)
        {
            var collection = collections[index];
            if (string.Equals(BuildAreaKey(collection), expectedArea, StringComparison.Ordinal))
            {
                return collection;
            }
        }

        throw new InvalidOperationException("Conditional formatting collection with area '" + expectedArea + "' was not found.");
    }

    private static string BuildAreaKey(FormatConditionCollection collection)
    {
        var areas = new List<string>(collection.RangeCount);
        for (var index = 0; index < collection.RangeCount; index++)
        {
            var area = collection.GetCellArea(index);
            var startCell = BuildCellName(area.FirstRow, area.FirstColumn);
            if (area.TotalRows == 1 && area.TotalColumns == 1)
            {
                areas.Add(startCell);
                continue;
            }

            var endCell = BuildCellName(area.FirstRow + area.TotalRows - 1, area.FirstColumn + area.TotalColumns - 1);
            areas.Add(startCell + ":" + endCell);
        }

        return string.Join(" ", areas);
    }

    private static string BuildCellName(int rowIndex, int columnIndex)
    {
        var dividend = columnIndex + 1;
        var name = string.Empty;
        while (dividend > 0)
        {
            var remainder = (dividend - 1) % 26;
            name = (char)('A' + remainder) + name;
            dividend = (dividend - remainder - 1) / 26;
        }

        return name + (rowIndex + 1).ToString();
    }

    private static void AssertArea(CellArea area, int firstRow, int firstColumn, int totalRows, int totalColumns)
    {
        AssertEx.Equal(firstRow, area.FirstRow);
        AssertEx.Equal(firstColumn, area.FirstColumn);
        AssertEx.Equal(totalRows, area.TotalRows);
        AssertEx.Equal(totalColumns, area.TotalColumns);
    }

    public static Workbook CreateAdvancedConditionalFormattingWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Advanced Conditional Formatting";

        sheet.Cells["A1"].PutValue("error one");
        sheet.Cells["A2"].PutValue("ok");
        sheet.Cells["B1"].PutValue("ok");
        sheet.Cells["B2"].PutValue("warning");
        sheet.Cells["C1"].PutValue("prefix-item");
        sheet.Cells["D1"].PutValue("item-suffix");
        sheet.Cells["E1"].PutValue(DateTime.UtcNow.Date);
        sheet.Cells["F1"].PutValue(1);
        sheet.Cells["F2"].PutValue(1);
        sheet.Cells["G1"].PutValue(10);
        sheet.Cells["G2"].PutValue(20);
        for (var index = 1; index <= 10; index++)
        {
            sheet.Cells["H" + index].PutValue(index * 10);
            sheet.Cells["I" + index].PutValue(index * 10);
            sheet.Cells["J" + index].PutValue(index * 10);
            sheet.Cells["K" + index].PutValue(index * 10);
            sheet.Cells["L" + index].PutValue(index * 10);
            sheet.Cells["M" + index].PutValue(index * 10);
            sheet.Cells["N" + index].PutValue(index * 10);
        }

        var contains = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        contains.AddArea(CellArea.CreateCellArea("A1", "A10"));
        var containsRule = contains[contains.AddCondition(FormatConditionType.ContainsText)];
        containsRule.Formula1 = "error";
        var containsStyle = containsRule.Style;
        containsStyle.Pattern = FillPattern.Solid;
        containsStyle.ForegroundColor = Color.FromArgb(255, 255, 235, 156);
        containsStyle.Font.Color = Color.FromArgb(255, 156, 0, 6);
        containsRule.Style = containsStyle;

        var notContains = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        notContains.AddArea(CellArea.CreateCellArea("B1", "B10"));
        var notContainsRule = notContains[notContains.AddCondition(FormatConditionType.NotContainsText)];
        notContainsRule.Formula1 = "warning";

        var beginsWith = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        beginsWith.AddArea(CellArea.CreateCellArea("C1", "C10"));
        var beginsWithRule = beginsWith[beginsWith.AddCondition(FormatConditionType.BeginsWith)];
        beginsWithRule.Formula1 = "prefix";

        var endsWith = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        endsWith.AddArea(CellArea.CreateCellArea("D1", "D10"));
        var endsWithRule = endsWith[endsWith.AddCondition(FormatConditionType.EndsWith)];
        endsWithRule.Formula1 = "suffix";

        var timePeriod = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        timePeriod.AddArea(CellArea.CreateCellArea("E1", "E10"));
        var timePeriodRule = timePeriod[timePeriod.AddCondition(FormatConditionType.TimePeriod)];
        timePeriodRule.TimePeriod = "today";
        var timeStyle = timePeriodRule.Style;
        timeStyle.Pattern = FillPattern.Solid;
        timeStyle.ForegroundColor = Color.FromArgb(255, 221, 235, 247);
        timePeriodRule.Style = timeStyle;

        var duplicate = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        duplicate.AddArea(CellArea.CreateCellArea("F1", "F10"));
        duplicate[duplicate.AddCondition(FormatConditionType.DuplicateValues)].Duplicate = true;

        var unique = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        unique.AddArea(CellArea.CreateCellArea("G1", "G10"));
        unique[unique.AddCondition(FormatConditionType.UniqueValues)].Duplicate = false;

        var top10 = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        top10.AddArea(CellArea.CreateCellArea("H1", "H10"));
        var top10Rule = top10[top10.AddCondition(FormatConditionType.Top10)];
        top10Rule.Top = true;
        top10Rule.Percent = true;
        top10Rule.Rank = 10;

        var bottom10 = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        bottom10.AddArea(CellArea.CreateCellArea("I1", "I10"));
        var bottom10Rule = bottom10[bottom10.AddCondition(FormatConditionType.Bottom10)];
        bottom10Rule.Top = false;
        bottom10Rule.Rank = 10;

        var aboveAverage = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        aboveAverage.AddArea(CellArea.CreateCellArea("J1", "J10"));
        var aboveAverageRule = aboveAverage[aboveAverage.AddCondition(FormatConditionType.AboveAverage)];
        aboveAverageRule.Above = true;

        var belowAverage = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        belowAverage.AddArea(CellArea.CreateCellArea("K1", "K10"));
        var belowAverageRule = belowAverage[belowAverage.AddCondition(FormatConditionType.BelowAverage)];
        belowAverageRule.Above = false;
        belowAverageRule.StandardDeviation = 1;

        var colorScale = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        colorScale.AddArea(CellArea.CreateCellArea("L1", "L10"));
        var colorScaleRule = colorScale[colorScale.AddCondition(FormatConditionType.ColorScale)];
        colorScaleRule.ColorScaleCount = 3;
        colorScaleRule.MinColor = Color.FromArgb(255, 248, 105, 107);
        colorScaleRule.MidColor = Color.FromArgb(255, 255, 235, 132);
        colorScaleRule.MaxColor = Color.FromArgb(255, 99, 190, 123);

        var dataBar = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        dataBar.AddArea(CellArea.CreateCellArea("M1", "M10"));
        var dataBarRule = dataBar[dataBar.AddCondition(FormatConditionType.DataBar)];
        dataBarRule.BarColor = Color.FromArgb(255, 99, 142, 198);
        dataBarRule.NegativeBarColor = Color.FromArgb(255, 255, 0, 0);
        dataBarRule.ShowBorder = true;
        dataBarRule.Direction = "left-to-right";

        var iconSet = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
        iconSet.AddArea(CellArea.CreateCellArea("N1", "N10"));
        var iconSetRule = iconSet[iconSet.AddCondition(FormatConditionType.IconSet)];
        iconSetRule.IconSetType = "4Arrows";
        iconSetRule.ReverseIcons = true;
        iconSetRule.ShowIconOnly = true;

        return workbook;
    }

    public static void AssertAdvancedConditionalFormattings(Workbook workbook)
    {
        var sheet = workbook.Worksheets["Advanced Conditional Formatting"];
        AssertEx.Equal(14, sheet.ConditionalFormattings.Count);

        AssertEx.Equal("error", GetCollectionByArea(sheet.ConditionalFormattings, "A1:A10")[0].Formula1);
        AssertEx.Equal(FormatConditionType.ContainsText, GetCollectionByArea(sheet.ConditionalFormattings, "A1:A10")[0].Type);
        AssertEx.Equal(FillPattern.Solid, GetCollectionByArea(sheet.ConditionalFormattings, "A1:A10")[0].Style.Pattern);

        AssertEx.Equal(FormatConditionType.NotContainsText, GetCollectionByArea(sheet.ConditionalFormattings, "B1:B10")[0].Type);
        AssertEx.Equal("warning", GetCollectionByArea(sheet.ConditionalFormattings, "B1:B10")[0].Formula1);

        AssertEx.Equal(FormatConditionType.BeginsWith, GetCollectionByArea(sheet.ConditionalFormattings, "C1:C10")[0].Type);
        AssertEx.Equal("prefix", GetCollectionByArea(sheet.ConditionalFormattings, "C1:C10")[0].Formula1);

        AssertEx.Equal(FormatConditionType.EndsWith, GetCollectionByArea(sheet.ConditionalFormattings, "D1:D10")[0].Type);
        AssertEx.Equal("suffix", GetCollectionByArea(sheet.ConditionalFormattings, "D1:D10")[0].Formula1);

        AssertEx.Equal(FormatConditionType.TimePeriod, GetCollectionByArea(sheet.ConditionalFormattings, "E1:E10")[0].Type);
        AssertEx.Equal("today", GetCollectionByArea(sheet.ConditionalFormattings, "E1:E10")[0].TimePeriod);

        AssertEx.Equal(FormatConditionType.DuplicateValues, GetCollectionByArea(sheet.ConditionalFormattings, "F1:F10")[0].Type);
        AssertEx.True(GetCollectionByArea(sheet.ConditionalFormattings, "F1:F10")[0].Duplicate);

        AssertEx.Equal(FormatConditionType.UniqueValues, GetCollectionByArea(sheet.ConditionalFormattings, "G1:G10")[0].Type);
        AssertEx.False(GetCollectionByArea(sheet.ConditionalFormattings, "G1:G10")[0].Duplicate);

        var topRule = GetCollectionByArea(sheet.ConditionalFormattings, "H1:H10")[0];
        AssertEx.Equal(FormatConditionType.Top10, topRule.Type);
        AssertEx.True(topRule.Top);
        AssertEx.True(topRule.Percent);
        AssertEx.Equal(10, topRule.Rank);

        var bottomRule = GetCollectionByArea(sheet.ConditionalFormattings, "I1:I10")[0];
        AssertEx.Equal(FormatConditionType.Bottom10, bottomRule.Type);
        AssertEx.False(bottomRule.Top);
        AssertEx.Equal(10, bottomRule.Rank);

        AssertEx.Equal(FormatConditionType.AboveAverage, GetCollectionByArea(sheet.ConditionalFormattings, "J1:J10")[0].Type);
        AssertEx.True(GetCollectionByArea(sheet.ConditionalFormattings, "J1:J10")[0].Above);

        var belowRule = GetCollectionByArea(sheet.ConditionalFormattings, "K1:K10")[0];
        AssertEx.Equal(FormatConditionType.BelowAverage, belowRule.Type);
        AssertEx.False(belowRule.Above);
        AssertEx.Equal(1, belowRule.StandardDeviation);

        var colorScaleRule = GetCollectionByArea(sheet.ConditionalFormattings, "L1:L10")[0];
        AssertEx.Equal(FormatConditionType.ColorScale, colorScaleRule.Type);
        AssertEx.Equal(3, colorScaleRule.ColorScaleCount);
        AssertEx.Equal(Color.FromArgb(255, 248, 105, 107), colorScaleRule.MinColor);
        AssertEx.Equal(Color.FromArgb(255, 255, 235, 132), colorScaleRule.MidColor);
        AssertEx.Equal(Color.FromArgb(255, 99, 190, 123), colorScaleRule.MaxColor);

        var dataBarRule = GetCollectionByArea(sheet.ConditionalFormattings, "M1:M10")[0];
        AssertEx.Equal(FormatConditionType.DataBar, dataBarRule.Type);
        AssertEx.Equal(Color.FromArgb(255, 99, 142, 198), dataBarRule.BarColor);

        var iconSetRule = GetCollectionByArea(sheet.ConditionalFormattings, "N1:N10")[0];
        AssertEx.Equal(FormatConditionType.IconSet, iconSetRule.Type);
        AssertEx.Equal("4Arrows", iconSetRule.IconSetType);
        AssertEx.True(iconSetRule.ReverseIcons);
        AssertEx.True(iconSetRule.ShowIconOnly);
    }}




public static class DefinedNameScenarioFactory
{
    public static Workbook CreateDefinedNamesWorkbook()
    {
        var workbook = PageSetupScenarioFactory.CreatePageSetupWorkbook();
        workbook.Worksheets.Add("Scoped Sheet");
        workbook.Worksheets[1].Cells["B2"].PutValue(5);

        var global = workbook.DefinedNames[workbook.DefinedNames.Add("GlobalRange", "='Print Sheet'!$A$1:$C$10")];
        global.Hidden = true;
        global.Comment = "Primary range";

        var local = workbook.DefinedNames[workbook.DefinedNames.Add("LocalCell", "'Scoped Sheet'!$B$2", 1)];
        local.Comment = "Scoped name";

        return workbook;
    }

    public static void AssertDefinedNames(Workbook workbook)
    {
        PageSetupScenarioFactory.AssertPageSetup(workbook);
        AssertEx.Equal(2, workbook.DefinedNames.Count);
        AssertEx.Equal("5", workbook.Worksheets[1].Cells["B2"].StringValue);

        var global = workbook.DefinedNames[0];
        AssertEx.Equal("GlobalRange", global.Name);
        AssertEx.Equal("'Print Sheet'!$A$1:$C$10", global.Formula);
        AssertEx.Null(global.LocalSheetIndex);
        AssertEx.True(global.Hidden);
        AssertEx.Equal("Primary range", global.Comment);

        var local = workbook.DefinedNames[1];
        AssertEx.Equal("LocalCell", local.Name);
        AssertEx.Equal("'Scoped Sheet'!$B$2", local.Formula);
        AssertEx.Equal(1, local.LocalSheetIndex ?? -1);
        AssertEx.False(local.Hidden);
        AssertEx.Equal("Scoped name", local.Comment);
    }
}

















