using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS;

internal static class XlsxWorkbookProperties
{
    internal static XElement? BuildWorkbookPropertiesElement(WorkbookModel model)
    {
        var element = new XElement(MainNs + "workbookPr");
        var hasState = false;
        if (model.Settings.DateSystem == Aspose.Cells_FOSS.Core.DateSystem.Mac1904)
        {
            element.SetAttributeValue("date1904", 1);
            hasState = true;
        }

        var properties = model.Properties;
        if (!string.IsNullOrEmpty(properties.CodeName))
        {
            element.SetAttributeValue("codeName", properties.CodeName);
            hasState = true;
        }

        if (!string.IsNullOrEmpty(properties.ShowObjects) && !string.Equals(properties.ShowObjects, "all", StringComparison.OrdinalIgnoreCase))
        {
            element.SetAttributeValue("showObjects", properties.ShowObjects);
            hasState = true;
        }

        if (properties.FilterPrivacy)
        {
            element.SetAttributeValue("filterPrivacy", 1);
            hasState = true;
        }

        if (!properties.ShowBorderUnselectedTables)
        {
            element.SetAttributeValue("showBorderUnselectedTables", 0);
            hasState = true;
        }

        if (!properties.ShowInkAnnotation)
        {
            element.SetAttributeValue("showInkAnnotation", 0);
            hasState = true;
        }

        if (properties.BackupFile)
        {
            element.SetAttributeValue("backupFile", 1);
            hasState = true;
        }

        if (!properties.SaveExternalLinkValues)
        {
            element.SetAttributeValue("saveExternalLinkValues", 0);
            hasState = true;
        }

        if (!string.IsNullOrEmpty(properties.UpdateLinks) && !string.Equals(properties.UpdateLinks, "userSet", StringComparison.OrdinalIgnoreCase))
        {
            element.SetAttributeValue("updateLinks", properties.UpdateLinks);
            hasState = true;
        }

        if (properties.HidePivotFieldList)
        {
            element.SetAttributeValue("hidePivotFieldList", 1);
            hasState = true;
        }

        if (properties.DefaultThemeVersion.HasValue)
        {
            element.SetAttributeValue("defaultThemeVersion", properties.DefaultThemeVersion.Value.ToString(CultureInfo.InvariantCulture));
            hasState = true;
        }

        return hasState ? element : null;
    }

    internal static XElement? BuildWorkbookProtectionElement(WorkbookModel model)
    {
        var protection = model.Properties.Protection;
        if (!protection.HasStoredState())
        {
            return null;
        }

        var element = new XElement(MainNs + "workbookProtection");
        if (protection.LockStructure)
        {
            element.SetAttributeValue("lockStructure", 1);
        }

        if (protection.LockWindows)
        {
            element.SetAttributeValue("lockWindows", 1);
        }

        if (protection.LockRevision)
        {
            element.SetAttributeValue("lockRevision", 1);
        }

        if (!string.IsNullOrEmpty(protection.WorkbookPassword))
        {
            element.SetAttributeValue("workbookPassword", protection.WorkbookPassword);
        }

        if (!string.IsNullOrEmpty(protection.RevisionsPassword))
        {
            element.SetAttributeValue("revisionsPassword", protection.RevisionsPassword);
        }

        return element;
    }

    internal static XElement? BuildBookViewsElement(WorkbookModel model)
    {
        var view = model.Properties.View;
        if (!view.HasStoredState(model.ActiveSheetIndex))
        {
            return null;
        }

        var workbookView = new XElement(MainNs + "workbookView");
        if (view.XWindow.HasValue)
        {
            workbookView.SetAttributeValue("xWindow", view.XWindow.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (view.YWindow.HasValue)
        {
            workbookView.SetAttributeValue("yWindow", view.YWindow.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (view.WindowWidth.HasValue)
        {
            workbookView.SetAttributeValue("windowWidth", view.WindowWidth.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (view.WindowHeight.HasValue)
        {
            workbookView.SetAttributeValue("windowHeight", view.WindowHeight.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (model.ActiveSheetIndex > 0 && model.ActiveSheetIndex < model.Worksheets.Count)
        {
            workbookView.SetAttributeValue("activeTab", model.ActiveSheetIndex.ToString(CultureInfo.InvariantCulture));
        }

        if (view.FirstSheet.HasValue)
        {
            var firstSheet = view.FirstSheet.Value;
            if (model.Worksheets.Count > 0 && firstSheet >= model.Worksheets.Count)
            {
                firstSheet = model.Worksheets.Count - 1;
            }

            if (firstSheet < 0)
            {
                firstSheet = 0;
            }

            workbookView.SetAttributeValue("firstSheet", firstSheet.ToString(CultureInfo.InvariantCulture));
        }

        if (view.ShowHorizontalScroll.HasValue)
        {
            workbookView.SetAttributeValue("showHorizontalScroll", view.ShowHorizontalScroll.Value ? 1 : 0);
        }

        if (view.ShowVerticalScroll.HasValue)
        {
            workbookView.SetAttributeValue("showVerticalScroll", view.ShowVerticalScroll.Value ? 1 : 0);
        }

        if (view.ShowSheetTabs.HasValue)
        {
            workbookView.SetAttributeValue("showSheetTabs", view.ShowSheetTabs.Value ? 1 : 0);
        }

        if (view.TabRatio.HasValue)
        {
            workbookView.SetAttributeValue("tabRatio", view.TabRatio.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (!string.IsNullOrEmpty(view.Visibility) && !string.Equals(view.Visibility, "visible", StringComparison.OrdinalIgnoreCase))
        {
            workbookView.SetAttributeValue("visibility", view.Visibility);
        }

        if (view.Minimized)
        {
            workbookView.SetAttributeValue("minimized", 1);
        }

        if (!view.AutoFilterDateGrouping)
        {
            workbookView.SetAttributeValue("autoFilterDateGrouping", 0);
        }

        return new XElement(MainNs + "bookViews", workbookView);
    }

    internal static XElement? BuildCalculationPropertiesElement(WorkbookModel model)
    {
        var calculation = model.Properties.Calculation;
        if (!calculation.HasStoredState())
        {
            return null;
        }

        var element = new XElement(MainNs + "calcPr");
        if (calculation.CalculationId.HasValue)
        {
            element.SetAttributeValue("calcId", calculation.CalculationId.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (!string.IsNullOrEmpty(calculation.CalculationMode) && !string.Equals(calculation.CalculationMode, "auto", StringComparison.OrdinalIgnoreCase))
        {
            element.SetAttributeValue("calcMode", calculation.CalculationMode);
        }

        if (calculation.FullCalculationOnLoad)
        {
            element.SetAttributeValue("fullCalcOnLoad", 1);
        }

        if (!string.IsNullOrEmpty(calculation.ReferenceMode) && !string.Equals(calculation.ReferenceMode, "A1", StringComparison.OrdinalIgnoreCase))
        {
            element.SetAttributeValue("refMode", calculation.ReferenceMode);
        }

        if (calculation.Iterate)
        {
            element.SetAttributeValue("iterate", 1);
        }

        if (calculation.IterateCount.HasValue)
        {
            element.SetAttributeValue("iterateCount", calculation.IterateCount.Value.ToString(CultureInfo.InvariantCulture));
        }

        if (calculation.IterateDelta.HasValue)
        {
            element.SetAttributeValue("iterateDelta", calculation.IterateDelta.Value.ToString("0.################", CultureInfo.InvariantCulture));
        }

        if (calculation.FullPrecision.HasValue)
        {
            element.SetAttributeValue("fullPrecision", calculation.FullPrecision.Value ? 1 : 0);
        }

        if (calculation.CalculationCompleted.HasValue)
        {
            element.SetAttributeValue("calcCompleted", calculation.CalculationCompleted.Value ? 1 : 0);
        }

        if (calculation.CalculationOnSave.HasValue)
        {
            element.SetAttributeValue("calcOnSave", calculation.CalculationOnSave.Value ? 1 : 0);
        }

        if (calculation.ConcurrentCalculation.HasValue)
        {
            element.SetAttributeValue("concurrentCalc", calculation.ConcurrentCalculation.Value ? 1 : 0);
        }

        if (calculation.ForceFullCalculation)
        {
            element.SetAttributeValue("forceFullCalc", 1);
        }

        return element;
    }

    internal static void LoadWorkbookMetadata(XElement workbookRoot, WorkbookModel workbookModel, int sheetCount, LoadDiagnostics diagnostics, LoadOptions options)
    {
        var workbookProperties = workbookModel.Properties;
        var workbookPr = workbookRoot.Element(MainNs + "workbookPr");
        if (workbookPr is null)
        {
            workbookModel.Settings.DateSystem = Aspose.Cells_FOSS.Core.DateSystem.Windows1900;
        }
        else
        {
            workbookModel.Settings.DateSystem = ReadBoolAttribute(workbookPr, "date1904", diagnostics, options, false, "/xl/workbook.xml") ? Aspose.Cells_FOSS.Core.DateSystem.Mac1904 : Aspose.Cells_FOSS.Core.DateSystem.Windows1900;
            workbookProperties.CodeName = ReadStringAttribute(workbookPr, "codeName");
            workbookProperties.ShowObjects = ReadChoiceAttribute(workbookPr, "showObjects", diagnostics, options, "/xl/workbook.xml", WorkbookPropertySupport.NormalizeShowObjects);
            workbookProperties.FilterPrivacy = ReadBoolAttribute(workbookPr, "filterPrivacy", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.ShowBorderUnselectedTables = ReadBoolAttribute(workbookPr, "showBorderUnselectedTables", diagnostics, options, true, "/xl/workbook.xml");
            workbookProperties.ShowInkAnnotation = ReadBoolAttribute(workbookPr, "showInkAnnotation", diagnostics, options, true, "/xl/workbook.xml");
            workbookProperties.BackupFile = ReadBoolAttribute(workbookPr, "backupFile", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.SaveExternalLinkValues = ReadBoolAttribute(workbookPr, "saveExternalLinkValues", diagnostics, options, true, "/xl/workbook.xml");
            workbookProperties.UpdateLinks = ReadChoiceAttribute(workbookPr, "updateLinks", diagnostics, options, "/xl/workbook.xml", WorkbookPropertySupport.NormalizeUpdateLinks);
            workbookProperties.HidePivotFieldList = ReadBoolAttribute(workbookPr, "hidePivotFieldList", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.DefaultThemeVersion = ReadNonNegativeIntAttribute(workbookPr, "defaultThemeVersion", diagnostics, options, "/xl/workbook.xml");
        }

        var protection = workbookRoot.Element(MainNs + "workbookProtection");
        if (protection is not null)
        {
            workbookProperties.Protection.LockStructure = ReadBoolAttribute(protection, "lockStructure", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.Protection.LockWindows = ReadBoolAttribute(protection, "lockWindows", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.Protection.LockRevision = ReadBoolAttribute(protection, "lockRevision", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.Protection.WorkbookPassword = ReadStringAttribute(protection, "workbookPassword");
            workbookProperties.Protection.RevisionsPassword = ReadStringAttribute(protection, "revisionsPassword");
        }

        var workbookView = workbookRoot.Element(MainNs + "bookViews")?.Element(MainNs + "workbookView");
        if (workbookView is not null)
        {
            workbookProperties.View.XWindow = ReadNonNegativeIntAttribute(workbookView, "xWindow", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.View.YWindow = ReadNonNegativeIntAttribute(workbookView, "yWindow", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.View.WindowWidth = ReadNonNegativeIntAttribute(workbookView, "windowWidth", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.View.WindowHeight = ReadNonNegativeIntAttribute(workbookView, "windowHeight", diagnostics, options, "/xl/workbook.xml");
            var firstSheet = ReadNonNegativeIntAttribute(workbookView, "firstSheet", diagnostics, options, "/xl/workbook.xml");
            if (firstSheet.HasValue)
            {
                if (firstSheet.Value >= sheetCount)
                {
                    AddWorkbookMetadataIssue(diagnostics, options, "/xl/workbook.xml", "Workbook firstSheet exceeded the worksheet count and was clamped.");
                    firstSheet = sheetCount > 0 ? sheetCount - 1 : 0;
                }

                workbookProperties.View.FirstSheet = firstSheet;
            }

            workbookProperties.View.ShowHorizontalScroll = ReadNullableBoolAttribute(workbookView, "showHorizontalScroll", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.View.ShowVerticalScroll = ReadNullableBoolAttribute(workbookView, "showVerticalScroll", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.View.ShowSheetTabs = ReadNullableBoolAttribute(workbookView, "showSheetTabs", diagnostics, options, "/xl/workbook.xml");
            var tabRatio = ReadNonNegativeIntAttribute(workbookView, "tabRatio", diagnostics, options, "/xl/workbook.xml");
            if (tabRatio.HasValue)
            {
                if (tabRatio.Value > 1000)
                {
                    AddWorkbookMetadataIssue(diagnostics, options, "/xl/workbook.xml", "Workbook tabRatio was out of range and was ignored.");
                }
                else
                {
                    workbookProperties.View.TabRatio = tabRatio;
                }
            }

            workbookProperties.View.Visibility = ReadChoiceAttribute(workbookView, "visibility", diagnostics, options, "/xl/workbook.xml", WorkbookPropertySupport.NormalizeVisibility);
            workbookProperties.View.Minimized = ReadBoolAttribute(workbookView, "minimized", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.View.AutoFilterDateGrouping = ReadBoolAttribute(workbookView, "autoFilterDateGrouping", diagnostics, options, true, "/xl/workbook.xml");

            var activeTab = ReadNonNegativeIntAttribute(workbookView, "activeTab", diagnostics, options, "/xl/workbook.xml");
            if (activeTab.HasValue)
            {
                if (activeTab.Value >= sheetCount)
                {
                    AddWorkbookMetadataIssue(diagnostics, options, "/xl/workbook.xml", "Workbook activeTab exceeded the worksheet count and was ignored.");
                }
                else
                {
                    workbookModel.ActiveSheetIndex = activeTab.Value;
                }
            }
        }

        var calcPr = workbookRoot.Element(MainNs + "calcPr");
        if (calcPr is not null)
        {
            workbookProperties.Calculation.CalculationId = ReadNonNegativeIntAttribute(calcPr, "calcId", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.Calculation.CalculationMode = ReadChoiceAttribute(calcPr, "calcMode", diagnostics, options, "/xl/workbook.xml", WorkbookPropertySupport.NormalizeCalculationMode);
            workbookProperties.Calculation.FullCalculationOnLoad = ReadBoolAttribute(calcPr, "fullCalcOnLoad", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.Calculation.ReferenceMode = ReadChoiceAttribute(calcPr, "refMode", diagnostics, options, "/xl/workbook.xml", WorkbookPropertySupport.NormalizeReferenceMode);
            workbookProperties.Calculation.Iterate = ReadBoolAttribute(calcPr, "iterate", diagnostics, options, false, "/xl/workbook.xml");
            workbookProperties.Calculation.IterateCount = ReadNonNegativeIntAttribute(calcPr, "iterateCount", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.Calculation.IterateDelta = ReadNonNegativeDoubleAttribute(calcPr, "iterateDelta", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.Calculation.FullPrecision = ReadNullableBoolAttribute(calcPr, "fullPrecision", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.Calculation.CalculationCompleted = ReadNullableBoolAttribute(calcPr, "calcCompleted", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.Calculation.CalculationOnSave = ReadNullableBoolAttribute(calcPr, "calcOnSave", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.Calculation.ConcurrentCalculation = ReadNullableBoolAttribute(calcPr, "concurrentCalc", diagnostics, options, "/xl/workbook.xml");
            workbookProperties.Calculation.ForceFullCalculation = ReadBoolAttribute(calcPr, "forceFullCalc", diagnostics, options, false, "/xl/workbook.xml");
        }
    }

    private static string ReadStringAttribute(XElement element, string name)
    {
        return ((string?)element.Attribute(name) ?? string.Empty).Trim();
    }

    private static string ReadChoiceAttribute(XElement element, string attributeName, LoadDiagnostics diagnostics, LoadOptions options, string partUri, Func<string?, string> normalizer)
    {
        var attribute = element.Attribute(attributeName);
        if (attribute is null)
        {
            return string.Empty;
        }

        try
        {
            return normalizer(attribute.Value);
        }
        catch (CellsException)
        {
            AddWorkbookMetadataIssue(diagnostics, options, partUri, "Workbook metadata attribute '" + attributeName + "' had an invalid value and was ignored.");
            return string.Empty;
        }
    }

    private static bool ReadBoolAttribute(XElement element, string attributeName, LoadDiagnostics diagnostics, LoadOptions options, bool defaultValue, string partUri)
    {
        var value = ReadNullableBoolAttribute(element, attributeName, diagnostics, options, partUri);
        return value ?? defaultValue;
    }

    private static bool? ReadNullableBoolAttribute(XElement element, string attributeName, LoadDiagnostics diagnostics, LoadOptions options, string partUri)
    {
        var attribute = element.Attribute(attributeName);
        if (attribute is null)
        {
            return null;
        }

        if (TryReadBoolean(attribute.Value, out var value))
        {
            return value;
        }

        AddWorkbookMetadataIssue(diagnostics, options, partUri, "Workbook metadata attribute '" + attributeName + "' had an invalid Boolean value and was ignored.");
        return null;
    }

    private static int? ReadNonNegativeIntAttribute(XElement element, string attributeName, LoadDiagnostics diagnostics, LoadOptions options, string partUri)
    {
        var attribute = element.Attribute(attributeName);
        if (attribute is null)
        {
            return null;
        }

        if (int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) && value >= 0)
        {
            return value;
        }

        AddWorkbookMetadataIssue(diagnostics, options, partUri, "Workbook metadata attribute '" + attributeName + "' had an invalid integer value and was ignored.");
        return null;
    }

    private static double? ReadNonNegativeDoubleAttribute(XElement element, string attributeName, LoadDiagnostics diagnostics, LoadOptions options, string partUri)
    {
        var attribute = element.Attribute(attributeName);
        if (attribute is null)
        {
            return null;
        }

        if (double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) && value >= 0d)
        {
            return value;
        }

        AddWorkbookMetadataIssue(diagnostics, options, partUri, "Workbook metadata attribute '" + attributeName + "' had an invalid numeric value and was ignored.");
        return null;
    }

    private static bool TryReadBoolean(string rawValue, out bool value)
    {
        if (rawValue == "1" || string.Equals(rawValue, "true", StringComparison.OrdinalIgnoreCase))
        {
            value = true;
            return true;
        }

        if (rawValue == "0" || string.Equals(rawValue, "false", StringComparison.OrdinalIgnoreCase))
        {
            value = false;
            return true;
        }

        value = false;
        return false;
    }

    private static void AddWorkbookMetadataIssue(LoadDiagnostics diagnostics, LoadOptions options, string partUri, string message)
    {
        AddIssue(diagnostics, options, new LoadIssue("WB-L003", DiagnosticSeverity.Warning, message)
        {
            PartUri = partUri,
        });
    }
}
