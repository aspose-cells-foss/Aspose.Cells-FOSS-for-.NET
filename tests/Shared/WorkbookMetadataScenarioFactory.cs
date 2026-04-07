using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Testing;

public static class WorkbookMetadataScenarioFactory
{
    public static Workbook CreateWorkbookMetadataWorkbook()
    {
        var workbook = new Workbook();
        workbook.Settings.Date1904 = true;

        workbook.Worksheets[0].Name = "Summary";
        var dataIndex = workbook.Worksheets.Add("Data");
        var archiveIndex = workbook.Worksheets.Add("Archive");
        workbook.Worksheets[archiveIndex].VisibilityType = VisibilityType.Hidden;
        workbook.Worksheets[dataIndex].Cells[0, 0].PutValue("Ready");

        workbook.Properties.CodeName = "WorkbookCode";
        workbook.Properties.ShowObjects = "placeholders";
        workbook.Properties.FilterPrivacy = true;
        workbook.Properties.ShowBorderUnselectedTables = false;
        workbook.Properties.ShowInkAnnotation = false;
        workbook.Properties.BackupFile = true;
        workbook.Properties.SaveExternalLinkValues = false;
        workbook.Properties.UpdateLinks = "never";
        workbook.Properties.HidePivotFieldList = true;
        workbook.Properties.DefaultThemeVersion = 166925;

        workbook.Properties.Protection.LockStructure = true;
        workbook.Properties.Protection.LockWindows = true;
        workbook.Properties.Protection.WorkbookPassword = "ABCD";
        workbook.Properties.Protection.RevisionsPassword = "EFGH";

        workbook.Properties.View.XWindow = 120;
        workbook.Properties.View.YWindow = 240;
        workbook.Properties.View.WindowWidth = 16000;
        workbook.Properties.View.WindowHeight = 9000;
        workbook.Properties.View.ActiveTab = 1;
        workbook.Properties.View.FirstSheet = 1;
        workbook.Properties.View.ShowHorizontalScroll = false;
        workbook.Properties.View.ShowVerticalScroll = false;
        workbook.Properties.View.ShowSheetTabs = false;
        workbook.Properties.View.TabRatio = 700;
        workbook.Properties.View.Visibility = "visible";
        workbook.Properties.View.Minimized = false;
        workbook.Properties.View.AutoFilterDateGrouping = false;

        workbook.Properties.Calculation.CalculationId = 191029;
        workbook.Properties.Calculation.CalculationMode = "manual";
        workbook.Properties.Calculation.FullCalculationOnLoad = true;
        workbook.Properties.Calculation.ReferenceMode = "R1C1";
        workbook.Properties.Calculation.Iterate = true;
        workbook.Properties.Calculation.IterateCount = 9;
        workbook.Properties.Calculation.IterateDelta = 0.00125d;
        workbook.Properties.Calculation.FullPrecision = false;
        workbook.Properties.Calculation.CalculationCompleted = false;
        workbook.Properties.Calculation.CalculationOnSave = false;
        workbook.Properties.Calculation.ConcurrentCalculation = false;
        workbook.Properties.Calculation.ForceFullCalculation = true;

        workbook.DocumentProperties.Title = "Quarterly Summary";
        workbook.DocumentProperties.Subject = "Operations";
        workbook.DocumentProperties.Author = "Automation";
        workbook.DocumentProperties.Keywords = "finance;ops";
        workbook.DocumentProperties.Comments = "Generated during tests.";
        workbook.DocumentProperties.Category = "Reports";
        workbook.DocumentProperties.Company = "Aspose Cells FOSS";
        workbook.DocumentProperties.Manager = "Release";
        workbook.DocumentProperties.Core.LastModifiedBy = "Verifier";
        workbook.DocumentProperties.Core.Revision = "7";
        workbook.DocumentProperties.Core.ContentStatus = "Draft";
        workbook.DocumentProperties.Core.Created = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        workbook.DocumentProperties.Core.Modified = new DateTime(2024, 6, 7, 8, 9, 10, DateTimeKind.Utc);
        workbook.DocumentProperties.Extended.Application = "Aspose.Cells_FOSS Tests";
        workbook.DocumentProperties.Extended.AppVersion = "0.2";
        workbook.DocumentProperties.Extended.DocSecurity = 2;
        workbook.DocumentProperties.Extended.HyperlinkBase = "https://example.com/base/";
        workbook.DocumentProperties.Extended.ScaleCrop = true;
        workbook.DocumentProperties.Extended.LinksUpToDate = true;
        workbook.DocumentProperties.Extended.SharedDoc = true;

        return workbook;
    }

    public static void AssertWorkbookMetadata(Workbook workbook)
    {
        AssertEx.True(workbook.Settings.Date1904);
        AssertEx.Equal("WorkbookCode", workbook.Properties.CodeName);
        AssertEx.Equal("placeholders", workbook.Properties.ShowObjects);
        AssertEx.True(workbook.Properties.FilterPrivacy);
        AssertEx.False(workbook.Properties.ShowBorderUnselectedTables);
        AssertEx.False(workbook.Properties.ShowInkAnnotation);
        AssertEx.True(workbook.Properties.BackupFile);
        AssertEx.False(workbook.Properties.SaveExternalLinkValues);
        AssertEx.Equal("never", workbook.Properties.UpdateLinks);
        AssertEx.True(workbook.Properties.HidePivotFieldList);
        AssertEx.Equal(166925, workbook.Properties.DefaultThemeVersion ?? 0);

        AssertEx.True(workbook.Properties.Protection.IsProtected);
        AssertEx.True(workbook.Properties.Protection.LockStructure);
        AssertEx.True(workbook.Properties.Protection.LockWindows);
        AssertEx.Equal("ABCD", workbook.Properties.Protection.WorkbookPassword);
        AssertEx.Equal("EFGH", workbook.Properties.Protection.RevisionsPassword);

        AssertEx.Equal(120, workbook.Properties.View.XWindow);
        AssertEx.Equal(240, workbook.Properties.View.YWindow);
        AssertEx.Equal(16000, workbook.Properties.View.WindowWidth);
        AssertEx.Equal(9000, workbook.Properties.View.WindowHeight);
        AssertEx.Equal(1, workbook.Properties.View.ActiveTab);
        AssertEx.Equal(1, workbook.Properties.View.FirstSheet);
        AssertEx.False(workbook.Properties.View.ShowHorizontalScroll);
        AssertEx.False(workbook.Properties.View.ShowVerticalScroll);
        AssertEx.False(workbook.Properties.View.ShowSheetTabs);
        AssertEx.Equal(700, workbook.Properties.View.TabRatio);
        AssertEx.Equal("visible", workbook.Properties.View.Visibility);
        AssertEx.False(workbook.Properties.View.Minimized);
        AssertEx.False(workbook.Properties.View.AutoFilterDateGrouping);

        AssertEx.Equal(191029, workbook.Properties.Calculation.CalculationId ?? 0);
        AssertEx.Equal("manual", workbook.Properties.Calculation.CalculationMode);
        AssertEx.True(workbook.Properties.Calculation.FullCalculationOnLoad);
        AssertEx.Equal("R1C1", workbook.Properties.Calculation.ReferenceMode);
        AssertEx.True(workbook.Properties.Calculation.Iterate);
        AssertEx.Equal(9, workbook.Properties.Calculation.IterateCount);
        AssertEx.Equal(0.00125d, workbook.Properties.Calculation.IterateDelta);
        AssertEx.False(workbook.Properties.Calculation.FullPrecision);
        AssertEx.False(workbook.Properties.Calculation.CalculationCompleted);
        AssertEx.False(workbook.Properties.Calculation.CalculationOnSave);
        AssertEx.False(workbook.Properties.Calculation.ConcurrentCalculation);
        AssertEx.True(workbook.Properties.Calculation.ForceFullCalculation);

        AssertEx.Equal("Quarterly Summary", workbook.DocumentProperties.Title);
        AssertEx.Equal("Operations", workbook.DocumentProperties.Subject);
        AssertEx.Equal("Automation", workbook.DocumentProperties.Author);
        AssertEx.Equal("finance;ops", workbook.DocumentProperties.Keywords);
        AssertEx.Equal("Generated during tests.", workbook.DocumentProperties.Comments);
        AssertEx.Equal("Reports", workbook.DocumentProperties.Category);
        AssertEx.Equal("Aspose Cells FOSS", workbook.DocumentProperties.Company);
        AssertEx.Equal("Release", workbook.DocumentProperties.Manager);
        AssertEx.Equal("Verifier", workbook.DocumentProperties.Core.LastModifiedBy);
        AssertEx.Equal("7", workbook.DocumentProperties.Core.Revision);
        AssertEx.Equal("Draft", workbook.DocumentProperties.Core.ContentStatus);
        AssertEx.Equal(new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc), workbook.DocumentProperties.Core.Created ?? DateTime.MinValue);
        AssertEx.Equal(new DateTime(2024, 6, 7, 8, 9, 10, DateTimeKind.Utc), workbook.DocumentProperties.Core.Modified ?? DateTime.MinValue);
        AssertEx.Equal("Aspose.Cells_FOSS Tests", workbook.DocumentProperties.Extended.Application);
        AssertEx.Equal("0.2", workbook.DocumentProperties.Extended.AppVersion);
        AssertEx.Equal(2, workbook.DocumentProperties.Extended.DocSecurity);
        AssertEx.Equal("https://example.com/base/", workbook.DocumentProperties.Extended.HyperlinkBase);
        AssertEx.True(workbook.DocumentProperties.Extended.ScaleCrop);
        AssertEx.True(workbook.DocumentProperties.Extended.LinksUpToDate);
        AssertEx.True(workbook.DocumentProperties.Extended.SharedDoc);
    }
}

