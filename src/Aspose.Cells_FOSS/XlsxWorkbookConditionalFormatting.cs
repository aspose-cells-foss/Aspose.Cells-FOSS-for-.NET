using System.IO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookConditionalFormatting
    {
        internal static List<XElement> BuildConditionalFormattingElements(WorksheetModel worksheet, StylesheetSaveContext stylesheet)
        {
            var ordered = GetOrderedConditionalFormattings(worksheet.ConditionalFormattings);
            var elements = new List<XElement>(ordered.Count);

            for (var index = 0; index < ordered.Count; index++)
            {
                var formatting = ordered[index];
                if (formatting.Areas.Count == 0 || formatting.Conditions.Count == 0)
                {
                    continue;
                }

                var element = new XElement(MainNs + "conditionalFormatting",
                    new XAttribute("sqref", BuildSqref(formatting.Areas)));

                for (var conditionIndex = 0; conditionIndex < formatting.Conditions.Count; conditionIndex++)
                {
                    element.Add(BuildConditionRule(formatting, formatting.Conditions[conditionIndex], conditionIndex, stylesheet));
                }

                elements.Add(element);
            }

            return elements;
        }

        internal static void LoadConditionalFormattings(WorksheetModel worksheetModel, XElement worksheetRoot, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            worksheetModel.ConditionalFormattings.Clear();

            foreach (var formattingElement in worksheetRoot.Elements(MainNs + "conditionalFormatting"))
            {
                var formatting = new ConditionalFormattingModel();
                var sqref = (string)formattingElement.Attribute("sqref");
                if (!TryLoadAreas(formatting, sqref, diagnostics, options, sheetName))
                {
                    continue;
                }

                foreach (var ruleElement in formattingElement.Elements(MainNs + "cfRule"))
                {
                    if (!TryLoadCondition(formatting, ruleElement, stylesheet, diagnostics, options, sheetName, sqref))
                    {
                        continue;
                    }
                }

                if (formatting.Conditions.Count == 0)
                {
                    AddIssue(diagnostics, options, new LoadIssue("CF-L003", DiagnosticSeverity.LossyRecoverable, "A conditional formatting entry without any supported rules was dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = sqref,
                    });
                    continue;
                }

                worksheetModel.ConditionalFormattings.Add(formatting);
            }
        }

        private static XElement BuildConditionRule(ConditionalFormattingModel formatting, FormatConditionModel condition, int conditionIndex, StylesheetSaveContext stylesheet)
        {
            var rule = new XElement(MainNs + "cfRule",
                new XAttribute("type", ToRuleTypeName(condition)),
                new XAttribute("priority", condition.Priority > 0 ? condition.Priority : conditionIndex + 1));

            var differentialStyleIndex = stylesheet.GetDifferentialStyleIndex(condition);
            if (differentialStyleIndex.HasValue)
            {
                rule.SetAttributeValue("dxfId", differentialStyleIndex.Value);
            }

            if (condition.StopIfTrue)
            {
                rule.SetAttributeValue("stopIfTrue", 1);
            }

            switch (condition.Type)
            {
                case FormatConditionType.CellValue:
                    var operatorName = ToOperatorName(condition.Operator);
                    if (!string.IsNullOrEmpty(operatorName))
                    {
                        rule.SetAttributeValue("operator", operatorName);
                    }

                    AddFormulaElement(rule, condition.Formula1);
                    AddFormulaElement(rule, condition.Formula2);
                    break;
                case FormatConditionType.Expression:
                    AddFormulaElement(rule, condition.Formula1);
                    break;
                case FormatConditionType.ContainsText:
                case FormatConditionType.NotContainsText:
                case FormatConditionType.BeginsWith:
                case FormatConditionType.EndsWith:
                    if (!string.IsNullOrEmpty(condition.Formula1))
                    {
                        rule.SetAttributeValue("text", condition.Formula1);
                        AddFormulaElement(rule, BuildTextRuleFormula(condition.Type, condition.Formula1, GetAnchorCellReference(formatting)));
                    }
                    break;
                case FormatConditionType.TimePeriod:
                    if (!string.IsNullOrEmpty(condition.TimePeriod))
                    {
                        rule.SetAttributeValue("timePeriod", condition.TimePeriod);
                    }
                    break;
                case FormatConditionType.Top10:
                case FormatConditionType.Bottom10:
                    rule.SetAttributeValue("bottom", condition.Type == FormatConditionType.Bottom10 || !condition.Top ? 1 : 0);
                    if (condition.Percent)
                    {
                        rule.SetAttributeValue("percent", 1);
                    }

                    if (condition.Rank > 0)
                    {
                        rule.SetAttributeValue("rank", condition.Rank);
                    }
                    break;
                case FormatConditionType.AboveAverage:
                case FormatConditionType.BelowAverage:
                    if (condition.Type == FormatConditionType.BelowAverage || !condition.Above)
                    {
                        rule.SetAttributeValue("aboveAverage", 0);
                    }

                    if (condition.StandardDeviation > 0)
                    {
                        rule.SetAttributeValue("stdDev", condition.StandardDeviation);
                    }
                    break;
                case FormatConditionType.ColorScale:
                    rule.Add(BuildColorScaleElement(condition));
                    break;
                case FormatConditionType.DataBar:
                    rule.Add(BuildDataBarElement(condition));
                    break;
                case FormatConditionType.IconSet:
                    rule.Add(BuildIconSetElement(condition));
                    break;
            }

            return rule;
        }

        private static bool TryLoadAreas(ConditionalFormattingModel formatting, string sqref, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sqref))
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("A conditional formatting entry is missing sqref.");
                }

                AddIssue(diagnostics, options, new LoadIssue("CF-L003", DiagnosticSeverity.LossyRecoverable, "A conditional formatting entry without any valid areas was dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                });
                return false;
            }

            var tokens = sqref.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var invalidAreaReported = false;
            for (var index = 0; index < tokens.Length; index++)
            {
                var token = tokens[index];
                MergeRegion region;
                if (!TryParseMergeReference(token, out region))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The conditional formatting sqref '" + token + "' is invalid.");
                    }

                    if (!invalidAreaReported)
                    {
                        AddIssue(diagnostics, options, new LoadIssue("CF-L001", DiagnosticSeverity.LossyRecoverable, "One or more conditional formatting areas were invalid and were dropped during load.", repairApplied: true, dataLossRisk: true)
                        {
                            SheetName = sheetName,
                            CellRef = token,
                        });
                        invalidAreaReported = true;
                    }
                    continue;
                }

                formatting.Areas.Add(new CellArea(region.FirstRow, region.FirstColumn, region.TotalRows, region.TotalColumns));
            }

            FormatConditionCollection.SortAreas(formatting.Areas);
            if (formatting.Areas.Count == 0)
            {
                AddIssue(diagnostics, options, new LoadIssue("CF-L003", DiagnosticSeverity.LossyRecoverable, "A conditional formatting entry without any valid areas was dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = sqref,
                });
                return false;
            }

            return true;
        }

        private static bool TryLoadCondition(ConditionalFormattingModel formatting, XElement ruleElement, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string cellRef)
        {
            FormatConditionType type;
            if (!TryParseRuleType(ruleElement, options, diagnostics, sheetName, cellRef, out type))
            {
                return false;
            }

            var condition = new FormatConditionModel
            {
                Type = type,
                Operator = ParseOperatorType((string)ruleElement.Attribute("operator")),
                StopIfTrue = ParseBoolAttribute(ruleElement.Attribute("stopIfTrue")),
                Priority = ParsePriority(ruleElement.Attribute("priority"), formatting, diagnostics, options, sheetName, cellRef),
                Style = ResolveDifferentialStyle(ruleElement.Attribute("dxfId"), stylesheet, diagnostics, options, sheetName, cellRef),
            };
            ApplyDefaults(condition);

            var formulas = new List<XElement>(ruleElement.Elements(MainNs + "formula"));
            switch (type)
            {
                case FormatConditionType.Expression:
                    condition.Formula1 = formulas.Count > 0 ? NormalizeFormula((string)formulas[0]) : null;
                    break;
                case FormatConditionType.CellValue:
                    condition.Formula1 = formulas.Count > 0 ? NormalizeFormula((string)formulas[0]) : null;
                    condition.Formula2 = formulas.Count > 1 ? NormalizeFormula((string)formulas[1]) : null;
                    break;
                case FormatConditionType.ContainsText:
                case FormatConditionType.NotContainsText:
                case FormatConditionType.BeginsWith:
                case FormatConditionType.EndsWith:
                    condition.Formula1 = NormalizeText((string)ruleElement.Attribute("text"));
                    break;
                case FormatConditionType.TimePeriod:
                    condition.TimePeriod = NormalizeText((string)ruleElement.Attribute("timePeriod"));
                    break;
                case FormatConditionType.Top10:
                case FormatConditionType.Bottom10:
                    condition.Top = type == FormatConditionType.Top10;
                    condition.Percent = ParseBoolAttribute(ruleElement.Attribute("percent"));
                    condition.Rank = ParseNonNegativeIntAttribute(ruleElement.Attribute("rank"));
                    break;
                case FormatConditionType.AboveAverage:
                case FormatConditionType.BelowAverage:
                    condition.Above = type == FormatConditionType.AboveAverage;
                    condition.StandardDeviation = ParseNonNegativeIntAttribute(ruleElement.Attribute("stdDev"));
                    break;
                case FormatConditionType.ColorScale:
                    LoadColorScale(condition, ruleElement.Element(MainNs + "colorScale"));
                    break;
                case FormatConditionType.DataBar:
                    LoadDataBar(condition, ruleElement.Element(MainNs + "dataBar"));
                    break;
                case FormatConditionType.IconSet:
                    LoadIconSet(condition, ruleElement.Element(MainNs + "iconSet"));
                    break;
                case FormatConditionType.DuplicateValues:
                    condition.Duplicate = true;
                    break;
                case FormatConditionType.UniqueValues:
                    condition.Duplicate = false;
                    break;
            }

            formatting.Conditions.Add(condition);
            return true;
        }

        private static bool TryParseRuleType(XElement ruleElement, LoadOptions options, LoadDiagnostics diagnostics, string sheetName, string cellRef, out FormatConditionType type)
        {
            type = FormatConditionType.CellValue;
            var typeText = NormalizeToken((string)ruleElement.Attribute("type"));
            switch (typeText)
            {
                case "cellis":
                    type = FormatConditionType.CellValue;
                    return true;
                case "expression":
                    type = FormatConditionType.Expression;
                    return true;
                case "containstext":
                    type = FormatConditionType.ContainsText;
                    return true;
                case "notcontainstext":
                    type = FormatConditionType.NotContainsText;
                    return true;
                case "beginswith":
                    type = FormatConditionType.BeginsWith;
                    return true;
                case "endswith":
                    type = FormatConditionType.EndsWith;
                    return true;
                case "timeperiod":
                    type = FormatConditionType.TimePeriod;
                    return true;
                case "duplicatevalues":
                    type = FormatConditionType.DuplicateValues;
                    return true;
                case "uniquevalues":
                    type = FormatConditionType.UniqueValues;
                    return true;
                case "top10":
                    type = ParseBoolAttribute(ruleElement.Attribute("bottom")) ? FormatConditionType.Bottom10 : FormatConditionType.Top10;
                    return true;
                case "aboveaverage":
                    type = ruleElement.Attribute("aboveAverage") != null && !ParseBoolAttribute(ruleElement.Attribute("aboveAverage"))
                        ? FormatConditionType.BelowAverage
                        : FormatConditionType.AboveAverage;
                    return true;
                case "colorscale":
                    type = FormatConditionType.ColorScale;
                    return true;
                case "databar":
                    type = FormatConditionType.DataBar;
                    return true;
                case "iconset":
                    type = FormatConditionType.IconSet;
                    return true;
                default:
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The conditional formatting rule type '" + typeText + "' is not supported.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("CF-L002", DiagnosticSeverity.LossyRecoverable, "An unsupported conditional formatting rule was dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = cellRef,
                    });
                    return false;
            }
        }

        private static void ApplyDefaults(FormatConditionModel condition)
        {
            switch (condition.Type)
            {
                case FormatConditionType.DuplicateValues:
                    condition.Duplicate = true;
                    break;
                case FormatConditionType.UniqueValues:
                    condition.Duplicate = false;
                    break;
                case FormatConditionType.Top10:
                    condition.Top = true;
                    if (condition.Rank == 0)
                    {
                        condition.Rank = 10;
                    }
                    break;
                case FormatConditionType.Bottom10:
                    condition.Top = false;
                    if (condition.Rank == 0)
                    {
                        condition.Rank = 10;
                    }
                    break;
                case FormatConditionType.AboveAverage:
                    condition.Above = true;
                    break;
                case FormatConditionType.BelowAverage:
                    condition.Above = false;
                    break;
                case FormatConditionType.ColorScale:
                    if (condition.ColorScaleCount == 0)
                    {
                        condition.ColorScaleCount = 2;
                    }
                    break;
                case FormatConditionType.DataBar:
                    if (IsEmptyColor(condition.BarColor))
                    {
                        condition.BarColor = new ColorValue(255, 99, 142, 198);
                    }
                    break;
                case FormatConditionType.IconSet:
                    if (string.IsNullOrEmpty(condition.IconSetType))
                    {
                        condition.IconSetType = "3TrafficLights1";
                    }
                    break;
            }
        }

        private static int ParsePriority(XAttribute attribute, ConditionalFormattingModel formatting, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string cellRef)
        {
            var parsed = ParseIntAttribute(attribute);
            if (parsed.HasValue && parsed.Value > 0)
            {
                return parsed.Value;
            }

            if (attribute != null)
            {
                AddIssue(diagnostics, options, new LoadIssue("CF-R001", DiagnosticSeverity.Recoverable, "A conditional formatting priority was invalid and was normalized during load.", repairApplied: true)
                {
                    SheetName = sheetName,
                    CellRef = cellRef,
                });
            }

            return formatting.Conditions.Count + 1;
        }

        private static StyleValue ResolveDifferentialStyle(XAttribute attribute, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string cellRef)
        {
            var dxfId = ParseIntAttribute(attribute);
            if (!dxfId.HasValue)
            {
                return StyleValue.Default.Clone();
            }

            if (dxfId.Value >= 0 && dxfId.Value < stylesheet.DifferentialFormats.Count)
            {
                return stylesheet.DifferentialFormats[dxfId.Value].Clone();
            }

            if (options.StrictMode)
            {
                throw new InvalidFileFormatException("The conditional formatting dxfId '" + dxfId.Value.ToString(CultureInfo.InvariantCulture) + "' is invalid.");
            }

            AddIssue(diagnostics, options, new LoadIssue("CF-R002", DiagnosticSeverity.Recoverable, "A conditional formatting differential style reference was invalid and the rule style was reset to default during load.", repairApplied: true)
            {
                SheetName = sheetName,
                CellRef = cellRef,
            });
            return StyleValue.Default.Clone();
        }

        private static void LoadColorScale(FormatConditionModel condition, XElement colorScaleElement)
        {
            if (colorScaleElement == null)
            {
                return;
            }

            var cfvoCount = 0;
            foreach (var _ in colorScaleElement.Elements(MainNs + "cfvo"))
            {
                cfvoCount++;
            }
            condition.ColorScaleCount = cfvoCount >= 3 ? 3 : 2;
            var colors = new List<XElement>(colorScaleElement.Elements(MainNs + "color"));
            if (colors.Count > 0)
            {
                condition.MinColor = ReadRgbColor(colors[0]);
            }

            if (condition.ColorScaleCount == 3 && colors.Count > 2)
            {
                condition.MidColor = ReadRgbColor(colors[1]);
                condition.MaxColor = ReadRgbColor(colors[2]);
                return;
            }

            if (colors.Count > 1)
            {
                condition.MaxColor = ReadRgbColor(colors[1]);
            }
        }

        private static void LoadDataBar(FormatConditionModel condition, XElement dataBarElement)
        {
            if (dataBarElement == null)
            {
                return;
            }

            var color = dataBarElement.Element(MainNs + "color");
            if (color != null)
            {
                condition.BarColor = ReadRgbColor(color);
            }
        }

        private static void LoadIconSet(FormatConditionModel condition, XElement iconSetElement)
        {
            if (iconSetElement == null)
            {
                return;
            }

            condition.IconSetType = NormalizeText((string)iconSetElement.Attribute("iconSet")) ?? "3TrafficLights1";
            condition.ReverseIcons = ParseBoolAttribute(iconSetElement.Attribute("reverse"));
            condition.ShowIconOnly = iconSetElement.Attribute("showValue") != null && !ParseBoolAttribute(iconSetElement.Attribute("showValue"));
        }

        private static XElement BuildColorScaleElement(FormatConditionModel condition)
        {
            var element = new XElement(MainNs + "colorScale",
                new XElement(MainNs + "cfvo", new XAttribute("type", "min")));

            if (condition.ColorScaleCount == 3)
            {
                element.Add(new XElement(MainNs + "cfvo",
                    new XAttribute("type", "percentile"),
                    new XAttribute("val", 50)));
            }

            element.Add(new XElement(MainNs + "cfvo", new XAttribute("type", "max")));
            element.Add(BuildColorElement(condition.MinColor, new ColorValue(255, 248, 105, 107)));
            if (condition.ColorScaleCount == 3)
            {
                element.Add(BuildColorElement(condition.MidColor, new ColorValue(255, 255, 235, 132)));
            }

            element.Add(BuildColorElement(condition.MaxColor, new ColorValue(255, 99, 190, 123)));
            return element;
        }

        private static XElement BuildDataBarElement(FormatConditionModel condition)
        {
            var element = new XElement(MainNs + "dataBar",
                new XElement(MainNs + "cfvo", new XAttribute("type", "min")),
                new XElement(MainNs + "cfvo", new XAttribute("type", "max")),
                BuildColorElement(condition.BarColor, new ColorValue(255, 99, 142, 198)));
            return element;
        }

        private static XElement BuildIconSetElement(FormatConditionModel condition)
        {
            var iconSetType = string.IsNullOrEmpty(condition.IconSetType) ? "3TrafficLights1" : condition.IconSetType;
            var element = new XElement(MainNs + "iconSet",
                new XAttribute("iconSet", iconSetType));
            if (condition.ReverseIcons)
            {
                element.SetAttributeValue("reverse", 1);
            }

            if (condition.ShowIconOnly)
            {
                element.SetAttributeValue("showValue", 0);
            }

            var iconCount = GetIconCount(iconSetType);
            for (var index = 0; index < iconCount; index++)
            {
                element.Add(new XElement(MainNs + "cfvo",
                    new XAttribute("type", "percent"),
                    new XAttribute("val", (100 * index) / iconCount)));
            }

            return element;
        }

        private static XElement BuildColorElement(ColorValue actual, ColorValue fallback)
        {
            var color = IsEmptyColor(actual) ? fallback : actual;
            return new XElement(MainNs + "color", new XAttribute("rgb", ToArgbHex(color)));
        }

        private static void AddFormulaElement(XElement parent, string formula)
        {
            if (!string.IsNullOrEmpty(formula))
            {
                parent.Add(new XElement(MainNs + "formula", formula));
            }
        }

        private static string BuildTextRuleFormula(FormatConditionType type, string text, string firstCell)
        {
            var escapedText = text.Replace("\"", "\"\"");
            switch (type)
            {
                case FormatConditionType.ContainsText:
                    return "NOT(ISERROR(SEARCH(\"" + escapedText + "\"," + firstCell + ")))";
                case FormatConditionType.NotContainsText:
                    return "ISERROR(SEARCH(\"" + escapedText + "\"," + firstCell + "))";
                case FormatConditionType.BeginsWith:
                    return "LEFT(" + firstCell + ",LEN(\"" + escapedText + "\"))=\"" + escapedText + "\"";
                case FormatConditionType.EndsWith:
                    return "RIGHT(" + firstCell + ",LEN(\"" + escapedText + "\"))=\"" + escapedText + "\"";
                default:
                    return string.Empty;
            }
        }

        private static string GetAnchorCellReference(ConditionalFormattingModel formatting)
        {
            if (formatting.Areas.Count == 0)
            {
                return "A1";
            }

            var area = formatting.Areas[0];
            return new CellAddress(area.FirstRow, area.FirstColumn).ToString();
        }

        private static List<ConditionalFormattingModel> GetOrderedConditionalFormattings(IReadOnlyList<ConditionalFormattingModel> collections)
        {
            var ordered = new List<ConditionalFormattingModel>(collections.Count);
            for (var index = 0; index < collections.Count; index++)
            {
                var collection = collections[index];
                if (collection.Areas.Count == 0 || collection.Conditions.Count == 0)
                {
                    continue;
                }

                FormatConditionCollection.SortAreas(collection.Areas);
                ordered.Add(collection);
            }

            ordered.Sort(CompareConditionalFormattings);
            return ordered;
        }

        private static int CompareConditionalFormattings(ConditionalFormattingModel left, ConditionalFormattingModel right)
        {
            return FormatConditionCollection.CompareAreas(left.Areas[0], right.Areas[0]);
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

        private static string ToRuleTypeName(FormatConditionModel condition)
        {
            switch (condition.Type)
            {
                case FormatConditionType.Expression:
                    return "expression";
                case FormatConditionType.ContainsText:
                    return "containsText";
                case FormatConditionType.NotContainsText:
                    return "notContainsText";
                case FormatConditionType.BeginsWith:
                    return "beginsWith";
                case FormatConditionType.EndsWith:
                    return "endsWith";
                case FormatConditionType.TimePeriod:
                    return "timePeriod";
                case FormatConditionType.DuplicateValues:
                    return "duplicateValues";
                case FormatConditionType.UniqueValues:
                    return "uniqueValues";
                case FormatConditionType.Top10:
                case FormatConditionType.Bottom10:
                    return "top10";
                case FormatConditionType.AboveAverage:
                case FormatConditionType.BelowAverage:
                    return "aboveAverage";
                case FormatConditionType.ColorScale:
                    return "colorScale";
                case FormatConditionType.DataBar:
                    return "dataBar";
                case FormatConditionType.IconSet:
                    return "iconSet";
                default:
                    return "cellIs";
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

        private static OperatorType ParseOperatorType(string text)
        {
            switch (NormalizeToken(text))
            {
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
                    return OperatorType.None;
            }
        }

        private static int ParseNonNegativeIntAttribute(XAttribute attribute)
        {
            var parsed = ParseIntAttribute(attribute);
            if (parsed.HasValue && parsed.Value >= 0)
            {
                return parsed.Value;
            }

            return 0;
        }

        private static int GetIconCount(string iconSetType)
        {
            if (string.IsNullOrEmpty(iconSetType))
            {
                return 3;
            }

            if (iconSetType[0] == '4')
            {
                return 4;
            }

            if (iconSetType[0] == '5')
            {
                return 5;
            }

            return 3;
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

        private static string NormalizeFormula(string value)
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
            if (value == null)
            {
                return null;
            }

            var trimmed = value.Trim();
            if (trimmed.Length == 0)
            {
                return null;
            }

            return trimmed;
        }

        private static ColorValue ReadRgbColor(XElement element)
        {
            var rgb = NormalizeText((string)element.Attribute("rgb"));
            if (string.IsNullOrEmpty(rgb) || rgb.Length != 8)
            {
                return default(ColorValue);
            }

            return new ColorValue(
                byte.Parse(rgb.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(rgb.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(rgb.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(rgb.Substring(6, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        }

        private static string ToArgbHex(ColorValue color)
        {
            return string.Concat(
                color.A.ToString("X2", CultureInfo.InvariantCulture),
                color.R.ToString("X2", CultureInfo.InvariantCulture),
                color.G.ToString("X2", CultureInfo.InvariantCulture),
                color.B.ToString("X2", CultureInfo.InvariantCulture));
        }

        private static bool IsEmptyColor(ColorValue color)
        {
            return !color.ThemeIndex.HasValue && !color.Indexed.HasValue
                && color.A == 0 && color.R == 0 && color.G == 0 && color.B == 0;
        }
    }
}
