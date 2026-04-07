using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS;

internal static class XlsxWorkbookPageSetup
{
    private const int MaxSpreadsheetRow = 1048575;
    private const int MaxSpreadsheetColumn = 16383;

    internal sealed class WorksheetDefinedNamesState
    {
        public string? PrintArea { get; set; }
        public string? PrintTitleRows { get; set; }
        public string? PrintTitleColumns { get; set; }
    }

    internal static XElement? BuildPageSetupDefinedNames(WorkbookModel model)
    {
        var definedNames = new List<XElement>();
        for (var sheetIndex = 0; sheetIndex < model.Worksheets.Count; sheetIndex++)
        {
            var worksheet = model.Worksheets[sheetIndex];
            var printArea = NormalizePrintAreaList(worksheet.PageSetup.PrintArea, worksheet.Name);
            if (!string.IsNullOrEmpty(printArea))
            {
                definedNames.Add(new XElement(MainNs + "definedName",
                    new XAttribute("name", DefinedNameUtility.PrintAreaDefinedName),
                    new XAttribute("localSheetId", sheetIndex),
                    printArea));
            }

            var printTitles = BuildPrintTitlesDefinedNameText(worksheet.PageSetup, worksheet.Name);
            if (!string.IsNullOrEmpty(printTitles))
            {
                definedNames.Add(new XElement(MainNs + "definedName",
                    new XAttribute("name", DefinedNameUtility.PrintTitlesDefinedName),
                    new XAttribute("localSheetId", sheetIndex),
                    printTitles));
            }
        }

        return definedNames.Count == 0 ? null : new XElement(MainNs + "definedNames", definedNames);
    }

    internal static Dictionary<int, WorksheetDefinedNamesState> LoadWorksheetDefinedNames(XElement workbookRoot, LoadDiagnostics diagnostics, LoadOptions options)
    {
        var states = new Dictionary<int, WorksheetDefinedNamesState>();
        foreach (var definedName in workbookRoot.Element(MainNs + "definedNames")?.Elements(MainNs + "definedName") ?? Enumerable.Empty<XElement>())
        {
            var name = ((string?)definedName.Attribute("name") ?? string.Empty).Trim();
            if (!string.Equals(name, DefinedNameUtility.PrintAreaDefinedName, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(name, DefinedNameUtility.PrintTitlesDefinedName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var localSheetId = ParseIntAttribute(definedName.Attribute("localSheetId"));
            if (!localSheetId.HasValue || localSheetId.Value < 0)
            {
                AddIssue(diagnostics, options, new LoadIssue("PG-L004", DiagnosticSeverity.Warning, $"Defined name '{name}' is missing a valid localSheetId and was ignored.", dataLossRisk: true));
                continue;
            }

            if (!states.TryGetValue(localSheetId.Value, out var state))
            {
                state = new WorksheetDefinedNamesState();
                states[localSheetId.Value] = state;
            }

            var text = definedName.Value?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                continue;
            }

            try
            {
                var resolvedState = state!;
                var definedNameText = text!;
                if (string.Equals(name, DefinedNameUtility.PrintAreaDefinedName, StringComparison.OrdinalIgnoreCase))
                {
                    resolvedState.PrintArea = NormalizeLoadedPrintArea(definedNameText);
                }
                else
                {
                    ParseLoadedPrintTitles(definedNameText, resolvedState);
                }
            }
            catch (CellsException exception)
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException($"Defined name '{name}' is invalid.", exception);
                }

                AddIssue(diagnostics, options, new LoadIssue("PG-L004", DiagnosticSeverity.Warning, $"Defined name '{name}' is invalid and was ignored.", dataLossRisk: true));
            }
        }

        return states;
    }

    internal static void ApplyWorksheetDefinedNames(WorksheetModel worksheetModel, WorksheetDefinedNamesState? definedNamesState)
    {
        if (definedNamesState is null)
        {
            return;
        }

        worksheetModel.PageSetup.PrintArea = definedNamesState.PrintArea;
        worksheetModel.PageSetup.PrintTitleRows = definedNamesState.PrintTitleRows;
        worksheetModel.PageSetup.PrintTitleColumns = definedNamesState.PrintTitleColumns;
    }

    internal static XElement? BuildSheetProperties(PageSetupModel pageSetup)
    {
        if (!pageSetup.FitToWidth.HasValue && !pageSetup.FitToHeight.HasValue)
        {
            return null;
        }

        return new XElement(MainNs + "sheetPr",
            new XElement(MainNs + "pageSetUpPr", new XAttribute("fitToPage", 1)));
    }

    internal static XElement? BuildPrintOptionsElement(PageSetupModel pageSetup)
    {
        if (!pageSetup.PrintOptions.GridLines
            && !pageSetup.PrintOptions.Headings
            && !pageSetup.PrintOptions.HorizontalCentered
            && !pageSetup.PrintOptions.VerticalCentered)
        {
            return null;
        }

        var element = new XElement(MainNs + "printOptions");
        if (pageSetup.PrintOptions.Headings)
        {
            element.SetAttributeValue("headings", 1);
        }

        if (pageSetup.PrintOptions.GridLines)
        {
            element.SetAttributeValue("gridLines", 1);
            element.SetAttributeValue("gridLinesSet", 1);
        }

        if (pageSetup.PrintOptions.HorizontalCentered)
        {
            element.SetAttributeValue("horizontalCentered", 1);
        }

        if (pageSetup.PrintOptions.VerticalCentered)
        {
            element.SetAttributeValue("verticalCentered", 1);
        }

        return element;
    }

    internal static XElement? BuildPageMarginsElement(PageSetupModel pageSetup)
    {
        if (MarginsEqual(pageSetup.Margins, new PageMarginsModel()))
        {
            return null;
        }

        return new XElement(MainNs + "pageMargins",
            new XAttribute("left", pageSetup.Margins.Left.ToString("R", CultureInfo.InvariantCulture)),
            new XAttribute("right", pageSetup.Margins.Right.ToString("R", CultureInfo.InvariantCulture)),
            new XAttribute("top", pageSetup.Margins.Top.ToString("R", CultureInfo.InvariantCulture)),
            new XAttribute("bottom", pageSetup.Margins.Bottom.ToString("R", CultureInfo.InvariantCulture)),
            new XAttribute("header", pageSetup.Margins.Header.ToString("R", CultureInfo.InvariantCulture)),
            new XAttribute("footer", pageSetup.Margins.Footer.ToString("R", CultureInfo.InvariantCulture)));
    }

    internal static XElement? BuildPageSetupElement(PageSetupModel pageSetup)
    {
        if (pageSetup.PaperSize == 0
            && pageSetup.Orientation == PageOrientation.Default
            && !pageSetup.FirstPageNumber.HasValue
            && !pageSetup.Scale.HasValue
            && !pageSetup.FitToWidth.HasValue
            && !pageSetup.FitToHeight.HasValue)
        {
            return null;
        }

        var element = new XElement(MainNs + "pageSetup");
        if (pageSetup.PaperSize > 0)
        {
            element.SetAttributeValue("paperSize", pageSetup.PaperSize);
        }

        if (pageSetup.Scale.HasValue)
        {
            element.SetAttributeValue("scale", pageSetup.Scale.Value);
        }

        if (pageSetup.FitToWidth.HasValue)
        {
            element.SetAttributeValue("fitToWidth", pageSetup.FitToWidth.Value);
        }

        if (pageSetup.FitToHeight.HasValue)
        {
            element.SetAttributeValue("fitToHeight", pageSetup.FitToHeight.Value);
        }

        if (pageSetup.FirstPageNumber.HasValue)
        {
            element.SetAttributeValue("firstPageNumber", pageSetup.FirstPageNumber.Value);
            element.SetAttributeValue("useFirstPageNumber", 1);
        }

        if (pageSetup.Orientation == PageOrientation.Portrait)
        {
            element.SetAttributeValue("orientation", "portrait");
        }
        else if (pageSetup.Orientation == PageOrientation.Landscape)
        {
            element.SetAttributeValue("orientation", "landscape");
        }

        return element;
    }

    internal static XElement? BuildHeaderFooterElement(PageSetupModel pageSetup)
    {
        var hasHeaderFooter = !string.IsNullOrEmpty(pageSetup.HeaderFooter.LeftHeader)
            || !string.IsNullOrEmpty(pageSetup.HeaderFooter.CenterHeader)
            || !string.IsNullOrEmpty(pageSetup.HeaderFooter.RightHeader)
            || !string.IsNullOrEmpty(pageSetup.HeaderFooter.LeftFooter)
            || !string.IsNullOrEmpty(pageSetup.HeaderFooter.CenterFooter)
            || !string.IsNullOrEmpty(pageSetup.HeaderFooter.RightFooter);
        if (!hasHeaderFooter)
        {
            return null;
        }

        var element = new XElement(MainNs + "headerFooter");
        if (!string.IsNullOrEmpty(pageSetup.HeaderFooter.LeftHeader) || !string.IsNullOrEmpty(pageSetup.HeaderFooter.CenterHeader) || !string.IsNullOrEmpty(pageSetup.HeaderFooter.RightHeader))
        {
            element.Add(new XElement(MainNs + "oddHeader", ComposeHeaderFooterText(pageSetup.HeaderFooter.LeftHeader, pageSetup.HeaderFooter.CenterHeader, pageSetup.HeaderFooter.RightHeader)));
        }

        if (!string.IsNullOrEmpty(pageSetup.HeaderFooter.LeftFooter) || !string.IsNullOrEmpty(pageSetup.HeaderFooter.CenterFooter) || !string.IsNullOrEmpty(pageSetup.HeaderFooter.RightFooter))
        {
            element.Add(new XElement(MainNs + "oddFooter", ComposeHeaderFooterText(pageSetup.HeaderFooter.LeftFooter, pageSetup.HeaderFooter.CenterFooter, pageSetup.HeaderFooter.RightFooter)));
        }

        return element;
    }

    internal static XElement? BuildRowBreaksElement(PageSetupModel pageSetup)
    {
        return BuildBreaksElement("rowBreaks", pageSetup.HorizontalPageBreaks, MaxSpreadsheetColumn);
    }

    internal static XElement? BuildColumnBreaksElement(PageSetupModel pageSetup)
    {
        return BuildBreaksElement("colBreaks", pageSetup.VerticalPageBreaks, MaxSpreadsheetRow);
    }

    private static XElement? BuildBreaksElement(string elementName, IEnumerable<int> breaks, int maxIndex)
    {
        var distinctBreaks = new HashSet<int>();
        foreach (var value in breaks)
        {
            distinctBreaks.Add(value);
        }

        var orderedBreaks = new List<int>(distinctBreaks);
        orderedBreaks.Sort();
        if (orderedBreaks.Count == 0)
        {
            return null;
        }

        var breakElements = new List<XElement>(orderedBreaks.Count);
        foreach (var value in orderedBreaks)
        {
            breakElements.Add(new XElement(MainNs + "brk",
                new XAttribute("id", value),
                new XAttribute("max", maxIndex),
                new XAttribute("man", 1)));
        }

        return new XElement(MainNs + elementName,
            new XAttribute("count", orderedBreaks.Count),
            new XAttribute("manualBreakCount", orderedBreaks.Count),
            breakElements);
    }

    internal static void LoadWorksheetPageSetup(WorksheetModel worksheetModel, XElement worksheetRoot, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
    {
        var pageSetup = worksheetModel.PageSetup;
        LoadPageMargins(pageSetup, worksheetRoot.Element(MainNs + "pageMargins"), diagnostics, options, sheetName);
        LoadPageSetupCore(pageSetup, worksheetRoot.Element(MainNs + "pageSetup"), diagnostics, options, sheetName);
        LoadPrintOptions(pageSetup, worksheetRoot.Element(MainNs + "printOptions"));
        LoadHeaderFooter(pageSetup, worksheetRoot.Element(MainNs + "headerFooter"));
        LoadBreaks(pageSetup.HorizontalPageBreaks, worksheetRoot.Element(MainNs + "rowBreaks"), diagnostics, options, sheetName, "row");
        LoadBreaks(pageSetup.VerticalPageBreaks, worksheetRoot.Element(MainNs + "colBreaks"), diagnostics, options, sheetName, "column");
    }

    private static void LoadPageMargins(PageSetupModel pageSetup, XElement? marginsElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
    {
        if (marginsElement is null)
        {
            return;
        }

        pageSetup.Margins.Left = ParseMarginAttribute(marginsElement.Attribute("left"), pageSetup.Margins.Left, diagnostics, options, sheetName, "left");
        pageSetup.Margins.Right = ParseMarginAttribute(marginsElement.Attribute("right"), pageSetup.Margins.Right, diagnostics, options, sheetName, "right");
        pageSetup.Margins.Top = ParseMarginAttribute(marginsElement.Attribute("top"), pageSetup.Margins.Top, diagnostics, options, sheetName, "top");
        pageSetup.Margins.Bottom = ParseMarginAttribute(marginsElement.Attribute("bottom"), pageSetup.Margins.Bottom, diagnostics, options, sheetName, "bottom");
        pageSetup.Margins.Header = ParseMarginAttribute(marginsElement.Attribute("header"), pageSetup.Margins.Header, diagnostics, options, sheetName, "header");
        pageSetup.Margins.Footer = ParseMarginAttribute(marginsElement.Attribute("footer"), pageSetup.Margins.Footer, diagnostics, options, sheetName, "footer");
    }

    private static void LoadPageSetupCore(PageSetupModel pageSetup, XElement? pageSetupElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
    {
        if (pageSetupElement is null)
        {
            return;
        }

        pageSetup.PaperSize = ParsePositiveIntAttribute(pageSetupElement.Attribute("paperSize"), diagnostics, options, sheetName, "paperSize") ?? 0;
        pageSetup.FirstPageNumber = ParsePositiveIntAttribute(pageSetupElement.Attribute("firstPageNumber"), diagnostics, options, sheetName, "firstPageNumber");
        pageSetup.Scale = ParseBoundedIntAttribute(pageSetupElement.Attribute("scale"), 10, 400, diagnostics, options, sheetName, "scale");
        pageSetup.FitToWidth = ParseNonNegativeIntAttribute(pageSetupElement.Attribute("fitToWidth"), diagnostics, options, sheetName, "fitToWidth");
        pageSetup.FitToHeight = ParseNonNegativeIntAttribute(pageSetupElement.Attribute("fitToHeight"), diagnostics, options, sheetName, "fitToHeight");
        pageSetup.Orientation = ParseOrientation((string?)pageSetupElement.Attribute("orientation"), diagnostics, options, sheetName);
    }

    private static void LoadPrintOptions(PageSetupModel pageSetup, XElement? printOptionsElement)
    {
        if (printOptionsElement is null)
        {
            return;
        }

        pageSetup.PrintOptions.Headings = ParseBoolAttribute(printOptionsElement.Attribute("headings"));
        pageSetup.PrintOptions.GridLines = ParseBoolAttribute(printOptionsElement.Attribute("gridLines"));
        pageSetup.PrintOptions.HorizontalCentered = ParseBoolAttribute(printOptionsElement.Attribute("horizontalCentered"));
        pageSetup.PrintOptions.VerticalCentered = ParseBoolAttribute(printOptionsElement.Attribute("verticalCentered"));
    }

    private static void LoadHeaderFooter(PageSetupModel pageSetup, XElement? headerFooterElement)
    {
        if (headerFooterElement is null)
        {
            return;
        }

        ParseHeaderFooterText((string?)headerFooterElement.Element(MainNs + "oddHeader"), out var leftHeader, out var centerHeader, out var rightHeader);
        pageSetup.HeaderFooter.LeftHeader = leftHeader;
        pageSetup.HeaderFooter.CenterHeader = centerHeader;
        pageSetup.HeaderFooter.RightHeader = rightHeader;

        ParseHeaderFooterText((string?)headerFooterElement.Element(MainNs + "oddFooter"), out var leftFooter, out var centerFooter, out var rightFooter);
        pageSetup.HeaderFooter.LeftFooter = leftFooter;
        pageSetup.HeaderFooter.CenterFooter = centerFooter;
        pageSetup.HeaderFooter.RightFooter = rightFooter;
    }

    private static void LoadBreaks(ICollection<int> target, XElement? breaksElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string axis)
    {
        target.Clear();
        if (breaksElement is null)
        {
            return;
        }

        foreach (var breakElement in breaksElement.Elements(MainNs + "brk"))
        {
            var id = ParseNonNegativeIntAttribute(breakElement.Attribute("id"), diagnostics, options, sheetName, axis + "Break");
            if (!id.HasValue)
            {
                continue;
            }

            if (!target.Contains(id.Value))
            {
                target.Add(id.Value);
            }
        }
    }

    private static string ComposeHeaderFooterText(string? left, string? center, string? right)
    {
        var builder = new StringBuilder();
        if (!string.IsNullOrEmpty(left))
        {
            builder.Append("&L").Append(left);
        }

        if (!string.IsNullOrEmpty(center))
        {
            builder.Append("&C").Append(center);
        }

        if (!string.IsNullOrEmpty(right))
        {
            builder.Append("&R").Append(right);
        }

        return builder.ToString();
    }

    private static void ParseHeaderFooterText(string? value, out string? left, out string? center, out string? right)
    {
        left = null;
        center = null;
        right = null;
        if (string.IsNullOrEmpty(value))
        {
            return;
        }

        var currentSection = 'C';
        var builder = new StringBuilder();

        var text = value!;
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] == '&' && index + 1 < text.Length)
            {
                var marker = char.ToUpperInvariant(text[index + 1]);
                if (marker == 'L' || marker == 'C' || marker == 'R')
                {
                    AssignHeaderFooterSection(currentSection, builder.ToString(), ref left, ref center, ref right);
                    builder.Clear();
                    currentSection = marker;
                    index++;
                    continue;
                }
            }

            builder.Append(text[index]);
        }

        AssignHeaderFooterSection(currentSection, builder.ToString(), ref left, ref center, ref right);
    }

    private static void AssignHeaderFooterSection(char section, string value, ref string? left, ref string? center, ref string? right)
    {
        var text = string.IsNullOrEmpty(value) ? null : value;
        if (section == 'L')
        {
            left = text;
        }
        else if (section == 'R')
        {
            right = text;
        }
        else
        {
            center = text;
        }
    }
    private static bool MarginsEqual(PageMarginsModel left, PageMarginsModel right)
    {
        return left.Left == right.Left
            && left.Right == right.Right
            && left.Top == right.Top
            && left.Bottom == right.Bottom
            && left.Header == right.Header
            && left.Footer == right.Footer;
    }

    private static string BuildPrintTitlesDefinedNameText(PageSetupModel pageSetup, string sheetName)
    {
        var segments = new List<string>();
        var normalizedRows = NormalizeTitleRows(pageSetup.PrintTitleRows, sheetName);
        if (!string.IsNullOrEmpty(normalizedRows))
        {
            segments.Add(normalizedRows);
        }

        var normalizedColumns = NormalizeTitleColumns(pageSetup.PrintTitleColumns, sheetName);
        if (!string.IsNullOrEmpty(normalizedColumns))
        {
            segments.Add(normalizedColumns);
        }

        return string.Join(",", segments);
    }

    private static string NormalizePrintAreaList(string? printArea, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(printArea))
        {
            return string.Empty;
        }

        var segments = new List<string>();
        var printAreaText = printArea!;
        foreach (var segment in SplitReferenceList(printAreaText))
        {
            segments.Add(QualifyReference(sheetName, NormalizeAreaReference(RemoveWorksheetQualifier(segment))));
        }

        return string.Join(",", segments);
    }

    private static string NormalizeLoadedPrintArea(string value)
    {
        var segments = new List<string>();
        foreach (var segment in SplitReferenceList(value))
        {
            segments.Add(NormalizeAreaReference(RemoveWorksheetQualifier(segment)));
        }

        return string.Join(",", segments);
    }

    private static void ParseLoadedPrintTitles(string value, WorksheetDefinedNamesState state)
    {
        foreach (var segment in SplitReferenceList(value))
        {
            var unqualified = RemoveWorksheetQualifier(segment);
            if (LooksLikeRowRange(unqualified))
            {
                state.PrintTitleRows = NormalizeRowReference(unqualified);
            }
            else if (LooksLikeColumnRange(unqualified))
            {
                state.PrintTitleColumns = NormalizeColumnReference(unqualified);
            }
            else
            {
                throw new CellsException("Print title reference is invalid.");
            }
        }
    }

    private static string NormalizeTitleRows(string? value, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return QualifyReference(sheetName, NormalizeRowReference(RemoveWorksheetQualifier(value!)));
    }

    private static string NormalizeTitleColumns(string? value, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return QualifyReference(sheetName, NormalizeColumnReference(RemoveWorksheetQualifier(value!)));
    }

    private static IEnumerable<string> SplitReferenceList(string value)
    {
        var parts = new List<string>();
        var builder = new StringBuilder();
        var inQuotes = false;

        var text = value!;
        for (var index = 0; index < text.Length; index++)
        {
            var character = value[index];
            if (character == '\'')
            {
                builder.Append(character);
                if (inQuotes && index + 1 < value.Length && value[index + 1] == '\'')
                {
                    builder.Append(value[index + 1]);
                    index++;
                    continue;
                }

                inQuotes = !inQuotes;
                continue;
            }

            if (character == ',' && !inQuotes)
            {
                var part = builder.ToString().Trim();
                if (part.Length > 0)
                {
                    parts.Add(part);
                }

                builder.Clear();
                continue;
            }

            builder.Append(character);
        }

        var lastPart = builder.ToString().Trim();
        if (lastPart.Length > 0)
        {
            parts.Add(lastPart);
        }

        return parts;
    }

    private static string RemoveWorksheetQualifier(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.Length == 0)
        {
            throw new CellsException("Worksheet reference is invalid.");
        }

        if (trimmed[0] == '\'')
        {
            var index = 1;
            while (index < trimmed.Length)
            {
                if (trimmed[index] == '\'')
                {
                    if (index + 1 < trimmed.Length && trimmed[index + 1] == '\'')
                    {
                        index += 2;
                        continue;
                    }

                    break;
                }

                index++;
            }

            if (index + 1 < trimmed.Length && trimmed[index + 1] == '!')
            {
                return trimmed.Substring(index + 2);
            }
        }

        var exclamation = trimmed.IndexOf('!');
        return exclamation >= 0 ? trimmed.Substring(exclamation + 1) : trimmed;
    }

    private static string QualifyReference(string sheetName, string reference)
    {
        return QuoteWorksheetName(sheetName) + "!" + reference;
    }

    private static string QuoteWorksheetName(string sheetName)
    {
        return "'" + sheetName.Replace("'", "''") + "'";
    }

    private static string NormalizeAreaReference(string value)
    {
        var parts = value.Split(':');
        if (parts.Length == 1)
        {
            var address = ParseAbsoluteCellAddress(parts[0]);
            return ToAbsoluteCellReference(address);
        }

        if (parts.Length != 2)
        {
            throw new CellsException("PrintArea must be a cell or range reference.");
        }

        var first = ParseAbsoluteCellAddress(parts[0]);
        var last = ParseAbsoluteCellAddress(parts[1]);
        if (last.RowIndex < first.RowIndex || last.ColumnIndex < first.ColumnIndex)
        {
            throw new CellsException("PrintArea range must be ordered from top-left to bottom-right.");
        }

        return ToAbsoluteCellReference(first) + ":" + ToAbsoluteCellReference(last);
    }

    private static string NormalizeRowReference(string value)
    {
        var normalized = value.Replace("$", string.Empty).Trim();
        var parts = normalized.Split(':');
        if (parts.Length != 2
            || !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var first)
            || !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var last)
            || first <= 0
            || last < first)
        {
            throw new CellsException("PrintTitleRows must be a row span like '1:2'.");
        }

        return "$" + first.ToString(CultureInfo.InvariantCulture) + ":$" + last.ToString(CultureInfo.InvariantCulture);
    }

    private static string NormalizeColumnReference(string value)
    {
        var normalized = value.Replace("$", string.Empty).Trim().ToUpperInvariant();
        var parts = normalized.Split(':');
        if (parts.Length != 2 || !IsColumnName(parts[0]) || !IsColumnName(parts[1]))
        {
            throw new CellsException("PrintTitleColumns must be a column span like 'A:B'.");
        }

        return "$" + parts[0] + ":$" + parts[1];
    }

    private static bool LooksLikeRowRange(string value)
    {
        return value.Replace("$", string.Empty).Split(':').All(delegate(string part) { return int.TryParse(part, NumberStyles.Integer, CultureInfo.InvariantCulture, out _); });
    }

    private static bool LooksLikeColumnRange(string value)
    {
        return value.Replace("$", string.Empty).Split(':').All(IsColumnName);
    }

    private static bool IsColumnName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        foreach (var character in value)
        {
            if (!char.IsLetter(character))
            {
                return false;
            }
        }

        return true;
    }

    private static CellAddress ParseAbsoluteCellAddress(string value)
    {
        var normalized = value.Replace("$", string.Empty).Trim();
        try
        {
            return CellAddress.Parse(normalized);
        }
        catch (ArgumentException exception)
        {
            throw new CellsException($"Cell reference '{value}' is invalid.", exception);
        }
    }

    private static string ToAbsoluteCellReference(CellAddress address)
    {
        var reference = address.ToString();
        var splitIndex = 0;
        while (splitIndex < reference.Length && char.IsLetter(reference[splitIndex]))
        {
            splitIndex++;
        }

        return "$" + reference.Substring(0, splitIndex) + "$" + reference.Substring(splitIndex);
    }

    private static double ParseMarginAttribute(XAttribute? attribute, double fallback, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string attributeName)
    {
        if (attribute is null)
        {
            return fallback;
        }

        var value = ParseDoubleAttribute(attribute);
        if (value.HasValue && value.Value >= 0d)
        {
            return value.Value;
        }

        if (options.StrictMode)
        {
            throw new InvalidFileFormatException($"The page margin attribute '{attributeName}' is invalid.");
        }

        AddIssue(diagnostics, options, new LoadIssue("PG-L001", DiagnosticSeverity.Warning, $"Page margin attribute '{attributeName}' is invalid and the default value was used.", dataLossRisk: true)
        {
            SheetName = sheetName,
        });
        return fallback;
    }

    private static int? ParsePositiveIntAttribute(XAttribute? attribute, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string attributeName)
    {
        if (attribute is null)
        {
            return null;
        }

        var value = ParseIntAttribute(attribute);
        if (value.HasValue && value.Value > 0)
        {
            return value.Value;
        }

        if (options.StrictMode)
        {
            throw new InvalidFileFormatException($"The page setup attribute '{attributeName}' is invalid.");
        }

        AddIssue(diagnostics, options, new LoadIssue("PG-L002", DiagnosticSeverity.Warning, $"Page setup attribute '{attributeName}' is invalid and was ignored.", dataLossRisk: true)
        {
            SheetName = sheetName,
        });
        return null;
    }

    private static int? ParseNonNegativeIntAttribute(XAttribute? attribute, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string attributeName)
    {
        if (attribute is null)
        {
            return null;
        }

        var value = ParseIntAttribute(attribute);
        if (value.HasValue && value.Value >= 0)
        {
            return value.Value;
        }

        if (options.StrictMode)
        {
            throw new InvalidFileFormatException($"The page setup attribute '{attributeName}' is invalid.");
        }

        AddIssue(diagnostics, options, new LoadIssue("PG-L002", DiagnosticSeverity.Warning, $"Page setup attribute '{attributeName}' is invalid and was ignored.", dataLossRisk: true)
        {
            SheetName = sheetName,
        });
        return null;
    }

    private static int? ParseBoundedIntAttribute(XAttribute? attribute, int minimum, int maximum, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string attributeName)
    {
        if (attribute is null)
        {
            return null;
        }

        var value = ParseIntAttribute(attribute);
        if (value.HasValue && value.Value >= minimum && value.Value <= maximum)
        {
            return value.Value;
        }

        if (options.StrictMode)
        {
            throw new InvalidFileFormatException($"The page setup attribute '{attributeName}' is invalid.");
        }

        AddIssue(diagnostics, options, new LoadIssue("PG-L002", DiagnosticSeverity.Warning, $"Page setup attribute '{attributeName}' is invalid and was ignored.", dataLossRisk: true)
        {
            SheetName = sheetName,
        });
        return null;
    }

    private static PageOrientation ParseOrientation(string? value, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return PageOrientation.Default;
        }

        if (string.Equals(value, "portrait", StringComparison.OrdinalIgnoreCase))
        {
            return PageOrientation.Portrait;
        }

        if (string.Equals(value, "landscape", StringComparison.OrdinalIgnoreCase))
        {
            return PageOrientation.Landscape;
        }

        if (options.StrictMode)
        {
            throw new InvalidFileFormatException("The page setup orientation is invalid.");
        }

        AddIssue(diagnostics, options, new LoadIssue("PG-L003", DiagnosticSeverity.Warning, "Page setup orientation is invalid and the default orientation was used.", dataLossRisk: true)
        {
            SheetName = sheetName,
        });
        return PageOrientation.Default;
    }
}








