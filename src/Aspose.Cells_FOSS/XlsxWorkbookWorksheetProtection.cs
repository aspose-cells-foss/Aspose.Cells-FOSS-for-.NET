using System.IO;
using System.Collections.Generic;
using System;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookWorksheetProtection
    {
        internal static XElement BuildSheetProtectionElement(WorksheetModel worksheet)
        {
            var protection = worksheet.Protection;
            if (!protection.HasStoredState())
            {
                return null;
            }

            var element = new XElement(MainNs + "sheetProtection");
            element.SetAttributeValue("sheet", 1);
            SetBoolAttributeWhenTrue(element, "objects", protection.Objects);
            SetBoolAttributeWhenTrue(element, "scenarios", protection.Scenarios);
            SetBoolAttributeWhenTrue(element, "formatCells", protection.FormatCells);
            SetBoolAttributeWhenTrue(element, "formatColumns", protection.FormatColumns);
            SetBoolAttributeWhenTrue(element, "formatRows", protection.FormatRows);
            SetBoolAttributeWhenTrue(element, "insertColumns", protection.InsertColumns);
            SetBoolAttributeWhenTrue(element, "insertRows", protection.InsertRows);
            SetBoolAttributeWhenTrue(element, "insertHyperlinks", protection.InsertHyperlinks);
            SetBoolAttributeWhenTrue(element, "deleteColumns", protection.DeleteColumns);
            SetBoolAttributeWhenTrue(element, "deleteRows", protection.DeleteRows);
            SetBoolAttributeWhenTrue(element, "selectLockedCells", protection.SelectLockedCells);
            SetBoolAttributeWhenTrue(element, "sort", protection.Sort);
            SetBoolAttributeWhenTrue(element, "autoFilter", protection.AutoFilter);
            SetBoolAttributeWhenTrue(element, "pivotTables", protection.PivotTables);
            SetBoolAttributeWhenTrue(element, "selectUnlockedCells", protection.SelectUnlockedCells);
            SetStringAttributeWhenPresent(element, "password", protection.PasswordHash);
            SetStringAttributeWhenPresent(element, "algorithmName", protection.AlgorithmName);
            SetStringAttributeWhenPresent(element, "hashValue", protection.HashValue);
            SetStringAttributeWhenPresent(element, "saltValue", protection.SaltValue);
            SetStringAttributeWhenPresent(element, "spinCount", protection.SpinCount);
            return element;
        }

        internal static void LoadWorksheetProtection(WorksheetModel worksheetModel, XElement worksheetRoot, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            worksheetModel.Protection.Clear();
            var protectionElement = worksheetRoot.Element(MainNs + "sheetProtection");
            if (protectionElement == null)
            {
                return;
            }

            worksheetModel.Protection.IsProtected = ParseProtectionBooleanAttribute(protectionElement.Attribute("sheet"), true, diagnostics, options, sheetName, "sheet");
            worksheetModel.Protection.Objects = ParseProtectionBooleanAttribute(protectionElement.Attribute("objects"), false, diagnostics, options, sheetName, "objects");
            worksheetModel.Protection.Scenarios = ParseProtectionBooleanAttribute(protectionElement.Attribute("scenarios"), false, diagnostics, options, sheetName, "scenarios");
            worksheetModel.Protection.FormatCells = ParseProtectionBooleanAttribute(protectionElement.Attribute("formatCells"), false, diagnostics, options, sheetName, "formatCells");
            worksheetModel.Protection.FormatColumns = ParseProtectionBooleanAttribute(protectionElement.Attribute("formatColumns"), false, diagnostics, options, sheetName, "formatColumns");
            worksheetModel.Protection.FormatRows = ParseProtectionBooleanAttribute(protectionElement.Attribute("formatRows"), false, diagnostics, options, sheetName, "formatRows");
            worksheetModel.Protection.InsertColumns = ParseProtectionBooleanAttribute(protectionElement.Attribute("insertColumns"), false, diagnostics, options, sheetName, "insertColumns");
            worksheetModel.Protection.InsertRows = ParseProtectionBooleanAttribute(protectionElement.Attribute("insertRows"), false, diagnostics, options, sheetName, "insertRows");
            worksheetModel.Protection.InsertHyperlinks = ParseProtectionBooleanAttribute(protectionElement.Attribute("insertHyperlinks"), false, diagnostics, options, sheetName, "insertHyperlinks");
            worksheetModel.Protection.DeleteColumns = ParseProtectionBooleanAttribute(protectionElement.Attribute("deleteColumns"), false, diagnostics, options, sheetName, "deleteColumns");
            worksheetModel.Protection.DeleteRows = ParseProtectionBooleanAttribute(protectionElement.Attribute("deleteRows"), false, diagnostics, options, sheetName, "deleteRows");
            worksheetModel.Protection.SelectLockedCells = ParseProtectionBooleanAttribute(protectionElement.Attribute("selectLockedCells"), false, diagnostics, options, sheetName, "selectLockedCells");
            worksheetModel.Protection.Sort = ParseProtectionBooleanAttribute(protectionElement.Attribute("sort"), false, diagnostics, options, sheetName, "sort");
            worksheetModel.Protection.AutoFilter = ParseProtectionBooleanAttribute(protectionElement.Attribute("autoFilter"), false, diagnostics, options, sheetName, "autoFilter");
            worksheetModel.Protection.PivotTables = ParseProtectionBooleanAttribute(protectionElement.Attribute("pivotTables"), false, diagnostics, options, sheetName, "pivotTables");
            worksheetModel.Protection.SelectUnlockedCells = ParseProtectionBooleanAttribute(protectionElement.Attribute("selectUnlockedCells"), false, diagnostics, options, sheetName, "selectUnlockedCells");
            worksheetModel.Protection.PasswordHash = ReadOptionalAttribute(protectionElement, "password");
            worksheetModel.Protection.AlgorithmName = ReadOptionalAttribute(protectionElement, "algorithmName");
            worksheetModel.Protection.HashValue = ReadOptionalAttribute(protectionElement, "hashValue");
            worksheetModel.Protection.SaltValue = ReadOptionalAttribute(protectionElement, "saltValue");
            worksheetModel.Protection.SpinCount = ReadOptionalAttribute(protectionElement, "spinCount");
        }

        private static bool ParseProtectionBooleanAttribute(XAttribute attribute, bool defaultValue, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string attributeName)
        {
            if (attribute == null)
            {
                return defaultValue;
            }

            var rawValue = ((string)attribute ?? string.Empty).Trim();
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
                throw new InvalidFileFormatException("Worksheet protection attribute '" + attributeName + "' is invalid.");
            }

            AddIssue(diagnostics, options, new LoadIssue("WS-L009", DiagnosticSeverity.Warning, "Worksheet protection attribute '" + attributeName + "' is invalid and the default value was used.", dataLossRisk: true)
            {
                SheetName = sheetName,
            });
            return defaultValue;
        }

        private static string ReadOptionalAttribute(XElement element, string attributeName)
        {
            var value = ((string)element.Attribute(attributeName) ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                return null;
            }

            return value;
        }

        private static void SetBoolAttributeWhenTrue(XElement element, string attributeName, bool value)
        {
            if (value)
            {
                element.SetAttributeValue(attributeName, 1);
            }
        }

        private static void SetStringAttributeWhenPresent(XElement element, string attributeName, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                element.SetAttributeValue(attributeName, value);
            }
        }
    }
}
