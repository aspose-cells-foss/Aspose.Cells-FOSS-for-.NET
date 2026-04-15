using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookPageSetup;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookDefinedNames
    {
        internal static XElement BuildDefinedNames(WorkbookModel model)
        {
            var definedNames = new List<XElement>();
            var pageSetupDefinedNames = BuildPageSetupDefinedNames(model);
            if (pageSetupDefinedNames != null)
            {
                foreach (var element in pageSetupDefinedNames.Elements())
                {
                    definedNames.Add(new XElement(element));
                }
            }

            var autoFilterDefinedNames = BuildAutoFilterDefinedNames(model);
            if (autoFilterDefinedNames != null)
            {
                foreach (var element in autoFilterDefinedNames.Elements())
                {
                    definedNames.Add(new XElement(element));
                }
            }

            for (var index = 0; index < model.DefinedNames.Count; index++)
            {
                var definedName = model.DefinedNames[index];
                var element = new XElement(MainNs + "definedName",
                    new XAttribute("name", definedName.Name),
                    definedName.Formula);

                if (definedName.LocalSheetIndex.HasValue)
                {
                    element.SetAttributeValue("localSheetId", definedName.LocalSheetIndex.Value.ToString(CultureInfo.InvariantCulture));
                }

                if (definedName.Hidden)
                {
                    element.SetAttributeValue("hidden", 1);
                }

                if (!string.IsNullOrEmpty(definedName.Comment))
                {
                    element.SetAttributeValue("comment", definedName.Comment);
                }

                definedNames.Add(element);
            }

            if (definedNames.Count == 0)
            {
                return null;
            }

            return new XElement(MainNs + "definedNames", definedNames);
        }

        internal static void LoadWorkbookDefinedNames(XElement workbookRoot, WorkbookModel workbookModel, int sheetCount, LoadDiagnostics diagnostics, LoadOptions options)
        {
            foreach (var element in workbookRoot.Element(MainNs + "definedNames")?.Elements(MainNs + "definedName") ?? Enumerable.Empty<XElement>())
            {
                var name = ((string)element.Attribute("name") ?? string.Empty).Trim();
                if (DefinedNameUtility.IsReservedName(name))
                {
                    continue;
                }

                if (name.Length == 0)
                {
                    HandleInvalidDefinedName(options, "Workbook defined name is missing a valid name.");
                    AddInvalidDefinedNameIssue(diagnostics, options, "Workbook defined name is missing a valid name and was ignored.");
                    continue;
                }

                int? localSheetIndex = null;
                var localSheetAttribute = element.Attribute("localSheetId");
                if (localSheetAttribute != null)
                {
                    localSheetIndex = ParseIntAttribute(localSheetAttribute);
                    if (!localSheetIndex.HasValue || localSheetIndex.Value < 0 || localSheetIndex.Value >= sheetCount)
                    {
                        HandleInvalidDefinedName(options, "Workbook defined name '" + name + "' has an invalid localSheetId.");
                        AddInvalidDefinedNameIssue(diagnostics, options, "Workbook defined name '" + name + "' has an invalid localSheetId and was ignored.");
                        continue;
                    }
                }

                string formula;
                try
                {
                    formula = DefinedNameUtility.NormalizeFormula(element.Value);
                }
                catch (CellsException exception)
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("Workbook defined name '" + name + "' is invalid.", exception);
                    }

                    AddInvalidDefinedNameIssue(diagnostics, options, "Workbook defined name '" + name + "' has an invalid formula and was ignored.");
                    continue;
                }

                if (ContainsDuplicate(workbookModel.DefinedNames, name, localSheetIndex))
                {
                    HandleInvalidDefinedName(options, "Workbook defined name '" + name + "' is duplicated in the same scope.");
                    AddInvalidDefinedNameIssue(diagnostics, options, "Workbook defined name '" + name + "' duplicates an existing scope and was ignored.");
                    continue;
                }

                workbookModel.DefinedNames.Add(new DefinedNameModel
                {
                    Name = name,
                    Formula = formula,
                    LocalSheetIndex = localSheetIndex,
                    Hidden = ParseBoolAttribute(element.Attribute("hidden")),
                    Comment = DefinedNameUtility.NormalizeComment((string)element.Attribute("comment")),
                });
            }
        }

        private static bool ContainsDuplicate(IReadOnlyList<DefinedNameModel> definedNames, string name, int? localSheetIndex)
        {
            for (var index = 0; index < definedNames.Count; index++)
            {
                var existing = definedNames[index];
                if (!string.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (DefinedNameUtility.SameScope(existing.LocalSheetIndex, localSheetIndex))
                {
                    return true;
                }
            }

            return false;
        }

        private static void AddInvalidDefinedNameIssue(LoadDiagnostics diagnostics, LoadOptions options, string message)
        {
            AddIssue(diagnostics, options, new LoadIssue("WB-L002", DiagnosticSeverity.Warning, message, dataLossRisk: true));
        }

        private static void HandleInvalidDefinedName(LoadOptions options, string message)
        {
            if (options.StrictMode)
            {
                throw new InvalidFileFormatException(message);
            }
        }

        private static XElement BuildAutoFilterDefinedNames(WorkbookModel model)
        {
            var definedNames = new List<XElement>();
            for (var sheetIndex = 0; sheetIndex < model.Worksheets.Count; sheetIndex++)
            {
                var worksheet = model.Worksheets[sheetIndex];
                if (string.IsNullOrEmpty(worksheet.AutoFilter.Range))
                {
                    continue;
                }

                definedNames.Add(new XElement(MainNs + "definedName",
                    new XAttribute("name", DefinedNameUtility.FilterDatabaseDefinedName),
                    new XAttribute("localSheetId", sheetIndex),
                    new XAttribute("hidden", 1),
                    QualifyAutoFilterRange(worksheet.Name, worksheet.AutoFilter.Range)));
            }

            return definedNames.Count == 0 ? null : new XElement(MainNs + "definedNames", definedNames);
        }

        private static string QualifyAutoFilterRange(string sheetName, string range)
        {
            var parts = range.Split(':');
            if (parts.Length == 1)
            {
                return QuoteWorksheetName(sheetName) + "!" + ToAbsoluteCellReference(parts[0]);
            }

            if (parts.Length != 2)
            {
                throw new CellsException("AutoFilter range is invalid.");
            }

            return QuoteWorksheetName(sheetName) + "!" + ToAbsoluteCellReference(parts[0]) + ":" + ToAbsoluteCellReference(parts[1]);
        }

        private static string QuoteWorksheetName(string sheetName)
        {
            return "'" + sheetName.Replace("'", "''") + "'";
        }

        private static string ToAbsoluteCellReference(string value)
        {
            CellAddress address;
            try
            {
                address = CellAddress.Parse(value);
            }
            catch (ArgumentException exception)
            {
                throw new CellsException("AutoFilter range is invalid.", exception);
            }

            var reference = address.ToString();
            var splitIndex = 0;
            while (splitIndex < reference.Length && char.IsLetter(reference[splitIndex]))
            {
                splitIndex++;
            }

            return "$" + reference.Substring(0, splitIndex) + "$" + reference.Substring(splitIndex);
        }
    }
}
