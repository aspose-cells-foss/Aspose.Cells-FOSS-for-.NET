using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS;

internal static class XlsxWorkbookWorksheetViews
{
    internal static XElement? BuildWorksheetSheetProperties(WorksheetModel worksheet)
    {
        var pageSetUpProperties = BuildPageSetupPropertiesElement(worksheet.PageSetup);
        var tabColorElement = BuildTabColorElement(worksheet.TabColor);
        if (pageSetUpProperties is null && tabColorElement is null)
        {
            return null;
        }

        var element = new XElement(MainNs + "sheetPr");
        if (tabColorElement is not null)
        {
            element.Add(tabColorElement);
        }

        if (pageSetUpProperties is not null)
        {
            element.Add(pageSetUpProperties);
        }

        return element;
    }

    internal static XElement? BuildWorksheetViewsElement(WorksheetModel worksheet)
    {
        if (WorksheetViewIsDefault(worksheet.View))
        {
            return null;
        }

        var sheetView = new XElement(MainNs + "sheetView",
            new XAttribute("workbookViewId", 0));

        if (!worksheet.View.ShowGridLines)
        {
            sheetView.SetAttributeValue("showGridLines", 0);
        }

        if (!worksheet.View.ShowRowColumnHeaders)
        {
            sheetView.SetAttributeValue("showRowColHeaders", 0);
        }

        if (!worksheet.View.ShowZeros)
        {
            sheetView.SetAttributeValue("showZeros", 0);
        }

        if (worksheet.View.RightToLeft)
        {
            sheetView.SetAttributeValue("rightToLeft", 1);
        }

        if (worksheet.View.ZoomScale != 100)
        {
            sheetView.SetAttributeValue("zoomScale", worksheet.View.ZoomScale);
        }

        return new XElement(MainNs + "sheetViews", sheetView);
    }

    internal static void LoadWorksheetViewSettings(WorksheetModel worksheetModel, XElement worksheetRoot, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
    {
        LoadTabColor(worksheetModel, worksheetRoot.Element(MainNs + "sheetPr"), diagnostics, options, sheetName);
        LoadSheetView(worksheetModel, worksheetRoot.Element(MainNs + "sheetViews")?.Element(MainNs + "sheetView"), diagnostics, options, sheetName);
    }

    private static XElement? BuildPageSetupPropertiesElement(PageSetupModel pageSetup)
    {
        if (!pageSetup.FitToWidth.HasValue && !pageSetup.FitToHeight.HasValue)
        {
            return null;
        }

        return new XElement(MainNs + "pageSetUpPr", new XAttribute("fitToPage", 1));
    }

    private static XElement? BuildTabColorElement(ColorValue? color)
    {
        if (!color.HasValue)
        {
            return null;
        }

        return new XElement(MainNs + "tabColor",
            new XAttribute("rgb", FormatColor(color.Value)));
    }

    private static bool WorksheetViewIsDefault(WorksheetViewModel view)
    {
        return view.ShowGridLines
            && view.ShowRowColumnHeaders
            && view.ShowZeros
            && !view.RightToLeft
            && view.ZoomScale == 100;
    }

    private static void LoadTabColor(WorksheetModel worksheetModel, XElement? sheetPropertiesElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
    {
        worksheetModel.TabColor = null;
        if (sheetPropertiesElement is null)
        {
            return;
        }

        var tabColor = sheetPropertiesElement.Element(MainNs + "tabColor");
        if (tabColor is null)
        {
            return;
        }

        var rgb = ((string?)tabColor.Attribute("rgb") ?? string.Empty).Trim();
        if (rgb.Length == 0)
        {
            return;
        }

        if (TryParseColor(rgb, out var color))
        {
            worksheetModel.TabColor = color;
            return;
        }

        if (options.StrictMode)
        {
            throw new InvalidFileFormatException("Worksheet tab color is invalid.");
        }

        AddIssue(diagnostics, options, new LoadIssue("WS-L007", DiagnosticSeverity.Warning, "Worksheet tab color is invalid and was ignored.", dataLossRisk: true)
        {
            SheetName = sheetName,
        });
    }

    private static void LoadSheetView(WorksheetModel worksheetModel, XElement? sheetViewElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
    {
        worksheetModel.View.ShowGridLines = true;
        worksheetModel.View.ShowRowColumnHeaders = true;
        worksheetModel.View.ShowZeros = true;
        worksheetModel.View.RightToLeft = false;
        worksheetModel.View.ZoomScale = 100;
        if (sheetViewElement is null)
        {
            return;
        }

        worksheetModel.View.ShowGridLines = ParseBooleanViewAttribute(sheetViewElement.Attribute("showGridLines"), true, diagnostics, options, sheetName, "showGridLines");
        worksheetModel.View.ShowRowColumnHeaders = ParseBooleanViewAttribute(sheetViewElement.Attribute("showRowColHeaders"), true, diagnostics, options, sheetName, "showRowColHeaders");
        worksheetModel.View.ShowZeros = ParseBooleanViewAttribute(sheetViewElement.Attribute("showZeros"), true, diagnostics, options, sheetName, "showZeros");
        worksheetModel.View.RightToLeft = ParseBooleanViewAttribute(sheetViewElement.Attribute("rightToLeft"), false, diagnostics, options, sheetName, "rightToLeft");
        worksheetModel.View.ZoomScale = ParseZoomScale(sheetViewElement.Attribute("zoomScale"), diagnostics, options, sheetName);
    }

    private static bool ParseBooleanViewAttribute(XAttribute? attribute, bool defaultValue, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string attributeName)
    {
        if (attribute is null)
        {
            return defaultValue;
        }

        var rawValue = ((string?)attribute ?? string.Empty).Trim();
        if (rawValue == "0" || string.Equals(rawValue, "false", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (rawValue == "1" || string.Equals(rawValue, "true", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (options.StrictMode)
        {
            throw new InvalidFileFormatException("Worksheet view attribute '" + attributeName + "' is invalid.");
        }

        AddIssue(diagnostics, options, new LoadIssue("WS-L008", DiagnosticSeverity.Warning, "Worksheet view attribute '" + attributeName + "' is invalid and the default value was used.", dataLossRisk: true)
        {
            SheetName = sheetName,
        });
        return defaultValue;
    }

    private static int ParseZoomScale(XAttribute? attribute, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
    {
        if (attribute is null)
        {
            return 100;
        }

        var value = ParseIntAttribute(attribute);
        if (value.HasValue && value.Value >= 10 && value.Value <= 400)
        {
            return value.Value;
        }

        if (options.StrictMode)
        {
            throw new InvalidFileFormatException("Worksheet zoomScale is invalid.");
        }

        AddIssue(diagnostics, options, new LoadIssue("WS-L008", DiagnosticSeverity.Warning, "Worksheet zoomScale is invalid and the default value was used.", dataLossRisk: true)
        {
            SheetName = sheetName,
        });
        return 100;
    }

    private static string FormatColor(ColorValue color)
    {
        return color.A.ToString("X2", CultureInfo.InvariantCulture)
            + color.R.ToString("X2", CultureInfo.InvariantCulture)
            + color.G.ToString("X2", CultureInfo.InvariantCulture)
            + color.B.ToString("X2", CultureInfo.InvariantCulture);
    }

    private static bool TryParseColor(string value, out ColorValue color)
    {
        color = default;
        if (value.Length != 8)
        {
            return false;
        }

        if (!byte.TryParse(value.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var a)
            || !byte.TryParse(value.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r)
            || !byte.TryParse(value.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g)
            || !byte.TryParse(value.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
        {
            return false;
        }

        color = new ColorValue(a, r, g, b);
        return true;
    }
}
