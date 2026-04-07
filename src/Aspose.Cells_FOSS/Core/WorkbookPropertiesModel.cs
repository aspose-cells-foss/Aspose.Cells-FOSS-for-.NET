using System;

namespace Aspose.Cells_FOSS.Core;

public sealed class WorkbookPropertiesModel
{
    public WorkbookPropertiesModel()
    {
        Protection = new WorkbookProtectionModel();
        View = new WorkbookViewModel();
        Calculation = new CalculationPropertiesModel();
    }

    public string CodeName { get; set; } = string.Empty;
    public string ShowObjects { get; set; } = string.Empty;
    public bool FilterPrivacy { get; set; }
    public bool ShowBorderUnselectedTables { get; set; } = true;
    public bool ShowInkAnnotation { get; set; } = true;
    public bool BackupFile { get; set; }
    public bool SaveExternalLinkValues { get; set; } = true;
    public string UpdateLinks { get; set; } = string.Empty;
    public bool HidePivotFieldList { get; set; }
    public int? DefaultThemeVersion { get; set; }
    public WorkbookProtectionModel Protection { get; }
    public WorkbookViewModel View { get; }
    public CalculationPropertiesModel Calculation { get; }

    public void CopyFrom(WorkbookPropertiesModel source)
    {
        CodeName = source.CodeName;
        ShowObjects = source.ShowObjects;
        FilterPrivacy = source.FilterPrivacy;
        ShowBorderUnselectedTables = source.ShowBorderUnselectedTables;
        ShowInkAnnotation = source.ShowInkAnnotation;
        BackupFile = source.BackupFile;
        SaveExternalLinkValues = source.SaveExternalLinkValues;
        UpdateLinks = source.UpdateLinks;
        HidePivotFieldList = source.HidePivotFieldList;
        DefaultThemeVersion = source.DefaultThemeVersion;
        Protection.CopyFrom(source.Protection);
        View.CopyFrom(source.View);
        Calculation.CopyFrom(source.Calculation);
    }

    public bool HasWorkbookPropertiesState()
    {
        return !string.IsNullOrEmpty(CodeName)
            || !string.IsNullOrEmpty(ShowObjects)
            || FilterPrivacy
            || !ShowBorderUnselectedTables
            || !ShowInkAnnotation
            || BackupFile
            || !SaveExternalLinkValues
            || !string.IsNullOrEmpty(UpdateLinks)
            || HidePivotFieldList
            || DefaultThemeVersion.HasValue;
    }
}

public sealed class WorkbookProtectionModel
{
    public bool LockStructure { get; set; }
    public bool LockWindows { get; set; }
    public bool LockRevision { get; set; }
    public string WorkbookPassword { get; set; } = string.Empty;
    public string RevisionsPassword { get; set; } = string.Empty;

    public void CopyFrom(WorkbookProtectionModel source)
    {
        LockStructure = source.LockStructure;
        LockWindows = source.LockWindows;
        LockRevision = source.LockRevision;
        WorkbookPassword = source.WorkbookPassword;
        RevisionsPassword = source.RevisionsPassword;
    }

    public bool HasStoredState()
    {
        return LockStructure
            || LockWindows
            || LockRevision
            || !string.IsNullOrEmpty(WorkbookPassword)
            || !string.IsNullOrEmpty(RevisionsPassword);
    }
}

public sealed class WorkbookViewModel
{
    public int? XWindow { get; set; }
    public int? YWindow { get; set; }
    public int? WindowWidth { get; set; }
    public int? WindowHeight { get; set; }
    public int? FirstSheet { get; set; }
    public bool? ShowHorizontalScroll { get; set; }
    public bool? ShowVerticalScroll { get; set; }
    public bool? ShowSheetTabs { get; set; }
    public int? TabRatio { get; set; }
    public string Visibility { get; set; } = string.Empty;
    public bool Minimized { get; set; }
    public bool AutoFilterDateGrouping { get; set; } = true;

    public void CopyFrom(WorkbookViewModel source)
    {
        XWindow = source.XWindow;
        YWindow = source.YWindow;
        WindowWidth = source.WindowWidth;
        WindowHeight = source.WindowHeight;
        FirstSheet = source.FirstSheet;
        ShowHorizontalScroll = source.ShowHorizontalScroll;
        ShowVerticalScroll = source.ShowVerticalScroll;
        ShowSheetTabs = source.ShowSheetTabs;
        TabRatio = source.TabRatio;
        Visibility = source.Visibility;
        Minimized = source.Minimized;
        AutoFilterDateGrouping = source.AutoFilterDateGrouping;
    }

    public bool HasStoredState(int activeSheetIndex)
    {
        return activeSheetIndex > 0
            || XWindow.HasValue
            || YWindow.HasValue
            || WindowWidth.HasValue
            || WindowHeight.HasValue
            || FirstSheet.HasValue
            || ShowHorizontalScroll.HasValue
            || ShowVerticalScroll.HasValue
            || ShowSheetTabs.HasValue
            || TabRatio.HasValue
            || !string.IsNullOrEmpty(Visibility)
            || Minimized
            || !AutoFilterDateGrouping;
    }
}

public sealed class CalculationPropertiesModel
{
    public int? CalculationId { get; set; }
    public string CalculationMode { get; set; } = string.Empty;
    public bool FullCalculationOnLoad { get; set; }
    public string ReferenceMode { get; set; } = string.Empty;
    public bool Iterate { get; set; }
    public int? IterateCount { get; set; }
    public double? IterateDelta { get; set; }
    public bool? FullPrecision { get; set; }
    public bool? CalculationCompleted { get; set; }
    public bool? CalculationOnSave { get; set; }
    public bool? ConcurrentCalculation { get; set; }
    public bool ForceFullCalculation { get; set; }

    public void CopyFrom(CalculationPropertiesModel source)
    {
        CalculationId = source.CalculationId;
        CalculationMode = source.CalculationMode;
        FullCalculationOnLoad = source.FullCalculationOnLoad;
        ReferenceMode = source.ReferenceMode;
        Iterate = source.Iterate;
        IterateCount = source.IterateCount;
        IterateDelta = source.IterateDelta;
        FullPrecision = source.FullPrecision;
        CalculationCompleted = source.CalculationCompleted;
        CalculationOnSave = source.CalculationOnSave;
        ConcurrentCalculation = source.ConcurrentCalculation;
        ForceFullCalculation = source.ForceFullCalculation;
    }

    public bool HasStoredState()
    {
        return CalculationId.HasValue
            || !string.IsNullOrEmpty(CalculationMode)
            || FullCalculationOnLoad
            || !string.IsNullOrEmpty(ReferenceMode)
            || Iterate
            || IterateCount.HasValue
            || IterateDelta.HasValue
            || FullPrecision.HasValue
            || CalculationCompleted.HasValue
            || CalculationOnSave.HasValue
            || ConcurrentCalculation.HasValue
            || ForceFullCalculation;
    }
}
