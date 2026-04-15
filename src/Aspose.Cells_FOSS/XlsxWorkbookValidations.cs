using System.IO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookValidations
    {
        internal static XElement BuildDataValidationsElement(WorksheetModel worksheet)
        {
            var ordered = GetOrderedValidations(worksheet.Validations);
            if (ordered.Count == 0)
            {
                return null;
            }

            var container = new XElement(MainNs + "dataValidations");
            var count = 0;

            for (var index = 0; index < ordered.Count; index++)
            {
                var validation = ordered[index];
                if (validation.Areas.Count == 0)
                {
                    continue;
                }

                var element = new XElement(MainNs + "dataValidation",
                    new XAttribute("sqref", BuildSqref(validation.Areas)));

                var typeName = ToValidationTypeName(validation.Type);
                if (!string.IsNullOrEmpty(typeName))
                {
                    element.SetAttributeValue("type", typeName);
                }

                var operatorName = ToOperatorName(validation.Operator);
                if (!string.IsNullOrEmpty(operatorName))
                {
                    element.SetAttributeValue("operator", operatorName);
                }

                if (validation.AlertStyle != ValidationAlertType.Stop)
                {
                    element.SetAttributeValue("errorStyle", ToAlertStyleName(validation.AlertStyle));
                }

                if (validation.IgnoreBlank)
                {
                    element.SetAttributeValue("allowBlank", 1);
                }

                if (!validation.InCellDropDown)
                {
                    element.SetAttributeValue("showDropDown", 1);
                }

                if (validation.ShowInput)
                {
                    element.SetAttributeValue("showInputMessage", 1);
                }

                if (validation.ShowError)
                {
                    element.SetAttributeValue("showErrorMessage", 1);
                }

                if (!string.IsNullOrEmpty(validation.ErrorTitle))
                {
                    element.SetAttributeValue("errorTitle", validation.ErrorTitle);
                }

                if (!string.IsNullOrEmpty(validation.ErrorMessage))
                {
                    element.SetAttributeValue("error", validation.ErrorMessage);
                }

                if (!string.IsNullOrEmpty(validation.InputTitle))
                {
                    element.SetAttributeValue("promptTitle", validation.InputTitle);
                }

                if (!string.IsNullOrEmpty(validation.InputMessage))
                {
                    element.SetAttributeValue("prompt", validation.InputMessage);
                }

                if (!string.IsNullOrEmpty(validation.Formula1))
                {
                    element.Add(new XElement(MainNs + "formula1", validation.Formula1));
                }

                if (!string.IsNullOrEmpty(validation.Formula2))
                {
                    element.Add(new XElement(MainNs + "formula2", validation.Formula2));
                }

                container.Add(element);
                count++;
            }

            if (count == 0)
            {
                return null;
            }

            container.SetAttributeValue("count", count.ToString(CultureInfo.InvariantCulture));
            return container;
        }

        internal static void LoadValidations(WorksheetModel worksheetModel, XElement worksheetRoot, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            worksheetModel.Validations.Clear();

            foreach (var validationElement in worksheetRoot.Element(MainNs + "dataValidations")?.Elements(MainNs + "dataValidation") ?? Enumerable.Empty<XElement>())
            {
                var validation = new ValidationModel();
                var sqref = (string)validationElement.Attribute("sqref");
                if (!TryLoadAreas(worksheetModel.Validations, validation, sqref, diagnostics, options, sheetName))
                {
                    continue;
                }

                validation.Type = ParseValidationType((string)validationElement.Attribute("type"), diagnostics, options, sheetName, sqref);
                validation.Operator = ParseOperatorType((string)validationElement.Attribute("operator"), diagnostics, options, sheetName, sqref);
                validation.AlertStyle = ParseAlertStyle((string)validationElement.Attribute("errorStyle"), diagnostics, options, sheetName, sqref);
                validation.IgnoreBlank = ParseBoolAttribute(validationElement.Attribute("allowBlank"));
                validation.InCellDropDown = !ParseBoolAttribute(validationElement.Attribute("showDropDown"));
                validation.ShowInput = ParseBoolAttribute(validationElement.Attribute("showInputMessage"));
                validation.ShowError = ParseBoolAttribute(validationElement.Attribute("showErrorMessage"));
                validation.ErrorTitle = NormalizeText((string)validationElement.Attribute("errorTitle"));
                validation.ErrorMessage = NormalizeText((string)validationElement.Attribute("error"));
                validation.InputTitle = NormalizeText((string)validationElement.Attribute("promptTitle"));
                validation.InputMessage = NormalizeText((string)validationElement.Attribute("prompt"));
                validation.Formula1 = NormalizeFormulaText((string)validationElement.Element(MainNs + "formula1"));
                validation.Formula2 = NormalizeFormulaText((string)validationElement.Element(MainNs + "formula2"));

                worksheetModel.Validations.Add(validation);
            }
        }

        private static bool TryLoadAreas(IReadOnlyList<ValidationModel> existingValidations, ValidationModel candidate, string sqref, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sqref))
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("A data validation is missing sqref.");
                }

                AddIssue(diagnostics, options, new LoadIssue("DV-L002", DiagnosticSeverity.LossyRecoverable, "A data validation without any valid areas was dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                });
                return false;
            }

            var tokens = sqref.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var invalidAreaReported = false;
            var overlapReported = false;

            for (var index = 0; index < tokens.Length; index++)
            {
                var token = tokens[index];
                MergeRegion region;
                if (!TryParseMergeReference(token, out region))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The data validation sqref '" + token + "' is invalid.");
                    }

                    if (!invalidAreaReported)
                    {
                        AddIssue(diagnostics, options, new LoadIssue("DV-L001", DiagnosticSeverity.LossyRecoverable, "One or more data validation areas were invalid and were dropped during load.", repairApplied: true, dataLossRisk: true)
                        {
                            SheetName = sheetName,
                            CellRef = token,
                        });
                        invalidAreaReported = true;
                    }
                    continue;
                }

                var area = new CellArea(region.FirstRow, region.FirstColumn, region.TotalRows, region.TotalColumns);
                if (ConflictsWithExisting(existingValidations, candidate, area))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The data validation sqref '" + token + "' overlaps an existing validation.");
                    }

                    if (!overlapReported)
                    {
                        AddIssue(diagnostics, options, new LoadIssue("DV-L003", DiagnosticSeverity.LossyRecoverable, "Overlapping data validation areas were normalized during load.", repairApplied: true, dataLossRisk: true)
                        {
                            SheetName = sheetName,
                            CellRef = token,
                        });
                        overlapReported = true;
                    }
                    continue;
                }

                candidate.Areas.Add(area);
            }

            ValidationCollection.SortAreas(candidate.Areas);
            if (candidate.Areas.Count == 0)
            {
                AddIssue(diagnostics, options, new LoadIssue("DV-L002", DiagnosticSeverity.LossyRecoverable, "A data validation without any valid areas was dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = sqref,
                });
                return false;
            }

            return true;
        }

        private static bool ConflictsWithExisting(IReadOnlyList<ValidationModel> existingValidations, ValidationModel candidate, CellArea area)
        {
            for (var validationIndex = 0; validationIndex < existingValidations.Count; validationIndex++)
            {
                var validation = existingValidations[validationIndex];
                for (var areaIndex = 0; areaIndex < validation.Areas.Count; areaIndex++)
                {
                    if (ValidationCollection.AreasOverlap(validation.Areas[areaIndex], area))
                    {
                        return true;
                    }
                }
            }

            for (var areaIndex = 0; areaIndex < candidate.Areas.Count; areaIndex++)
            {
                if (ValidationCollection.AreasOverlap(candidate.Areas[areaIndex], area))
                {
                    return true;
                }
            }

            return false;
        }

        private static List<ValidationModel> GetOrderedValidations(IReadOnlyList<ValidationModel> validations)
        {
            var ordered = new List<ValidationModel>(validations.Count);
            for (var index = 0; index < validations.Count; index++)
            {
                var validation = validations[index];
                if (validation.Areas.Count == 0)
                {
                    continue;
                }

                ValidationCollection.SortAreas(validation.Areas);
                ordered.Add(validation);
            }

            ordered.Sort(CompareValidations);
            return ordered;
        }

        private static int CompareValidations(ValidationModel left, ValidationModel right)
        {
            return ValidationCollection.CompareAreas(left.Areas[0], right.Areas[0]);
        }

        private static string BuildSqref(IReadOnlyList<CellArea> areas)
        {
            var references = new List<string>(areas.Count);
            for (var index = 0; index < areas.Count; index++)
            {
                references.Add(ToAreaReference(areas[index]));
            }

            return string.Join(" ", references);
        }

        private static string ToAreaReference(CellArea area)
        {
            var first = new CellAddress(area.FirstRow, area.FirstColumn).ToString();
            if (area.TotalRows == 1 && area.TotalColumns == 1)
            {
                return first;
            }

            var last = new CellAddress(area.FirstRow + area.TotalRows - 1, area.FirstColumn + area.TotalColumns - 1).ToString();
            return first + ":" + last;
        }

        private static ValidationType ParseValidationType(string text, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string cellRef)
        {
            switch (NormalizeToken(text))
            {
                case null:
                    return ValidationType.AnyValue;
                case "whole":
                    return ValidationType.WholeNumber;
                case "decimal":
                    return ValidationType.Decimal;
                case "list":
                    return ValidationType.List;
                case "date":
                    return ValidationType.Date;
                case "time":
                    return ValidationType.Time;
                case "textlength":
                    return ValidationType.TextLength;
                case "custom":
                    return ValidationType.Custom;
                default:
                    return WarnAndReturnAnyValue("type", text, diagnostics, options, sheetName, cellRef);
            }
        }

        private static OperatorType ParseOperatorType(string text, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string cellRef)
        {
            switch (NormalizeToken(text))
            {
                case null:
                    return OperatorType.None;
                case "between":
                    return OperatorType.Between;
                case "equal":
                    return OperatorType.Equal;
                case "greaterthan":
                    return OperatorType.GreaterThan;
                case "greaterthanorequal":
                    return OperatorType.GreaterOrEqual;
                case "lessthan":
                    return OperatorType.LessThan;
                case "lessthanorequal":
                    return OperatorType.LessOrEqual;
                case "notbetween":
                    return OperatorType.NotBetween;
                case "notequal":
                    return OperatorType.NotEqual;
                default:
                    return WarnAndReturnNone("operator", text, diagnostics, options, sheetName, cellRef);
            }
        }

        private static ValidationAlertType ParseAlertStyle(string text, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string cellRef)
        {
            switch (NormalizeToken(text))
            {
                case null:
                case "stop":
                    return ValidationAlertType.Stop;
                case "warning":
                    return ValidationAlertType.Warning;
                case "information":
                    return ValidationAlertType.Information;
                default:
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The data validation errorStyle '" + text + "' is invalid.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("DV-R001", DiagnosticSeverity.Recoverable, "An unknown data validation errorStyle was normalized during load.", repairApplied: true)
                    {
                        SheetName = sheetName,
                        CellRef = cellRef,
                    });
                    return ValidationAlertType.Stop;
            }
        }

        private static ValidationType WarnAndReturnAnyValue(string attributeName, string text, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string cellRef)
        {
            if (options.StrictMode)
            {
                throw new InvalidFileFormatException("The data validation " + attributeName + " '" + text + "' is invalid.");
            }

            AddIssue(diagnostics, options, new LoadIssue("DV-R001", DiagnosticSeverity.Recoverable, "An unknown data validation " + attributeName + " was normalized during load.", repairApplied: true)
            {
                SheetName = sheetName,
                CellRef = cellRef,
            });
            return ValidationType.AnyValue;
        }

        private static OperatorType WarnAndReturnNone(string attributeName, string text, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string cellRef)
        {
            if (options.StrictMode)
            {
                throw new InvalidFileFormatException("The data validation " + attributeName + " '" + text + "' is invalid.");
            }

            AddIssue(diagnostics, options, new LoadIssue("DV-R001", DiagnosticSeverity.Recoverable, "An unknown data validation " + attributeName + " was normalized during load.", repairApplied: true)
            {
                SheetName = sheetName,
                CellRef = cellRef,
            });
            return OperatorType.None;
        }

        private static string NormalizeToken(string value)
        {
            if (value == null)
            {
                return null;
            }

            var trimmed = value.Trim();
            if (trimmed.Length == 0)
            {
                return null;
            }

            return trimmed.ToLowerInvariant();
        }

        private static string NormalizeFormulaText(string value)
        {
            if (value == null)
            {
                return null;
            }

            var trimmed = value.Trim();
            if (trimmed.Length == 0)
            {
                return null;
            }

            if (trimmed[0] == '=')
            {
                return trimmed.Substring(1);
            }

            return trimmed;
        }

        private static string NormalizeText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }

            return value;
        }

        private static string ToValidationTypeName(ValidationType type)
        {
            switch (type)
            {
                case ValidationType.WholeNumber:
                    return "whole";
                case ValidationType.Decimal:
                    return "decimal";
                case ValidationType.List:
                    return "list";
                case ValidationType.Date:
                    return "date";
                case ValidationType.Time:
                    return "time";
                case ValidationType.TextLength:
                    return "textLength";
                case ValidationType.Custom:
                    return "custom";
                default:
                    return null;
            }
        }

        private static string ToOperatorName(OperatorType type)
        {
            switch (type)
            {
                case OperatorType.Between:
                    return "between";
                case OperatorType.Equal:
                    return "equal";
                case OperatorType.GreaterThan:
                    return "greaterThan";
                case OperatorType.GreaterOrEqual:
                    return "greaterThanOrEqual";
                case OperatorType.LessThan:
                    return "lessThan";
                case OperatorType.LessOrEqual:
                    return "lessThanOrEqual";
                case OperatorType.NotBetween:
                    return "notBetween";
                case OperatorType.NotEqual:
                    return "notEqual";
                default:
                    return null;
            }
        }

        private static string ToAlertStyleName(ValidationAlertType type)
        {
            switch (type)
            {
                case ValidationAlertType.Warning:
                    return "warning";
                case ValidationAlertType.Information:
                    return "information";
                default:
                    return "stop";
            }
        }
    }
}
