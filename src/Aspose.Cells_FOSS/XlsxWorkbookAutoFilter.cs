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
    internal static class XlsxWorkbookAutoFilter
    {
        internal static XElement BuildAutoFilterElement(WorksheetModel worksheet, int differentialFormatCount)
        {
            var autoFilter = worksheet.AutoFilter;
            if (!autoFilter.HasStoredState() || string.IsNullOrEmpty(autoFilter.Range))
            {
                return null;
            }

            var element = new XElement(MainNs + "autoFilter", new XAttribute("ref", autoFilter.Range));
            var orderedColumns = GetOrderedFilterColumns(autoFilter.FilterColumns);
            for (var index = 0; index < orderedColumns.Count; index++)
            {
                var filterColumn = orderedColumns[index];
                if (!filterColumn.HasStoredState())
                {
                    continue;
                }

                var filterColumnElement = BuildFilterColumnElement(filterColumn, differentialFormatCount);
                if (filterColumnElement != null)
                {
                    element.Add(filterColumnElement);
                }
            }

            var sortState = BuildSortStateElement(autoFilter.SortState, differentialFormatCount);
            if (sortState != null)
            {
                element.Add(sortState);
            }

            return element;
        }

        internal static void LoadAutoFilter(WorksheetModel worksheetModel, XElement worksheetRoot, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            worksheetModel.AutoFilter.Clear();
            var autoFilterElement = worksheetRoot.Element(MainNs + "autoFilter");
            if (autoFilterElement == null)
            {
                return;
            }

            var reference = (string)autoFilterElement.Attribute("ref");
            string normalizedRange;
            if (!AutoFilterSupport.TryNormalizeRange(reference, out normalizedRange))
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("The worksheet autoFilter ref '" + reference + "' is invalid.");
                }

                AddIssue(diagnostics, options, new LoadIssue("WS-L010", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter metadata had an invalid range and was dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = reference,
                });
                return;
            }

            worksheetModel.AutoFilter.Range = normalizedRange;

            var seenColumnIndexes = new HashSet<int>();
            foreach (var filterColumnElement in autoFilterElement.Elements(MainNs + "filterColumn"))
            {
                var column = LoadFilterColumn(filterColumnElement, stylesheet, diagnostics, options, sheetName, normalizedRange);
                if (column == null)
                {
                    continue;
                }

                if (!seenColumnIndexes.Add(column.ColumnIndex))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The worksheet autoFilter contains duplicate filter columns.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = normalizedRange,
                    });
                    continue;
                }

                worksheetModel.AutoFilter.FilterColumns.Add(column);
            }

            worksheetModel.AutoFilter.FilterColumns.Sort(delegate(FilterColumnModel left, FilterColumnModel right)
            {
                return AutoFilterSupport.CompareFilterColumns(left, right);
            });

            var sortStateElement = autoFilterElement.Element(MainNs + "sortState");
            if (sortStateElement != null)
            {
                LoadSortState(worksheetModel.AutoFilter.SortState, sortStateElement, stylesheet, diagnostics, options, sheetName, normalizedRange);
            }
        }

        private static XElement BuildFilterColumnElement(FilterColumnModel model, int differentialFormatCount)
        {
            var customFilters = BuildCustomFiltersElement(model);
            var filters = BuildFiltersElement(model.Filters);
            var colorFilter = BuildColorFilterElement(model.ColorFilter, differentialFormatCount);
            var dynamicFilter = BuildDynamicFilterElement(model.DynamicFilter);
            var top10 = BuildTop10Element(model.Top10);
            if (!model.HiddenButton
                && filters == null
                && customFilters == null
                && colorFilter == null
                && dynamicFilter == null
                && top10 == null)
            {
                return null;
            }

            var element = new XElement(MainNs + "filterColumn", new XAttribute("colId", model.ColumnIndex.ToString(CultureInfo.InvariantCulture)));
            if (model.HiddenButton)
            {
                element.SetAttributeValue("hiddenButton", 1);
            }

            if (colorFilter != null)
            {
                element.Add(colorFilter);
            }

            if (customFilters != null)
            {
                element.Add(customFilters);
            }

            if (dynamicFilter != null)
            {
                element.Add(dynamicFilter);
            }

            if (filters != null)
            {
                element.Add(filters);
            }

            if (top10 != null)
            {
                element.Add(top10);
            }

            return element;
        }

        private static XElement BuildFiltersElement(IReadOnlyList<string> filters)
        {
            if (filters.Count == 0)
            {
                return null;
            }

            var element = new XElement(MainNs + "filters");
            for (var index = 0; index < filters.Count; index++)
            {
                element.Add(new XElement(MainNs + "filter", new XAttribute("val", filters[index])));
            }

            return element;
        }

        private static XElement BuildCustomFiltersElement(FilterColumnModel model)
        {
            if (model.CustomFilters.Count == 0)
            {
                return null;
            }

            var element = new XElement(MainNs + "customFilters");
            if (model.CustomFiltersAnd)
            {
                element.SetAttributeValue("and", 1);
            }

            for (var index = 0; index < model.CustomFilters.Count; index++)
            {
                var filter = model.CustomFilters[index];
                var customFilter = new XElement(MainNs + "customFilter", new XAttribute("val", filter.Value));
                if (!string.IsNullOrEmpty(filter.Operator))
                {
                    customFilter.SetAttributeValue("operator", filter.Operator);
                }

                element.Add(customFilter);
            }

            return element;
        }

        private static XElement BuildColorFilterElement(AutoFilterColorFilterModel model, int differentialFormatCount)
        {
            if (!model.Enabled || !HasValidDifferentialStyleId(model.DifferentialStyleId, differentialFormatCount))
            {
                return null;
            }

            var element = new XElement(MainNs + "colorFilter");
            element.SetAttributeValue("dxfId", model.DifferentialStyleId.Value.ToString(CultureInfo.InvariantCulture));

            if (model.CellColor)
            {
                element.SetAttributeValue("cellColor", 1);
            }

            return element;
        }

        private static XElement BuildDynamicFilterElement(AutoFilterDynamicFilterModel model)
        {
            if (!model.Enabled)
            {
                return null;
            }

            var element = new XElement(MainNs + "dynamicFilter");
            if (!string.IsNullOrEmpty(model.Type))
            {
                element.SetAttributeValue("type", model.Type);
            }

            if (model.Value.HasValue)
            {
                element.SetAttributeValue("val", model.Value.Value.ToString("R", CultureInfo.InvariantCulture));
            }

            if (model.MaxValue.HasValue)
            {
                element.SetAttributeValue("maxVal", model.MaxValue.Value.ToString("R", CultureInfo.InvariantCulture));
            }

            return element;
        }

        private static XElement BuildTop10Element(AutoFilterTop10Model model)
        {
            if (!model.Enabled)
            {
                return null;
            }

            var element = new XElement(MainNs + "top10");
            if (!model.Top)
            {
                element.SetAttributeValue("top", 0);
            }

            if (model.Percent)
            {
                element.SetAttributeValue("percent", 1);
            }

            if (model.Value.HasValue)
            {
                element.SetAttributeValue("val", model.Value.Value.ToString("R", CultureInfo.InvariantCulture));
            }

            if (model.FilterValue.HasValue)
            {
                element.SetAttributeValue("filterVal", model.FilterValue.Value.ToString("R", CultureInfo.InvariantCulture));
            }

            return element;
        }

        private static XElement BuildSortStateElement(AutoFilterSortStateModel model, int differentialFormatCount)
        {
            if (!model.HasStoredState() || string.IsNullOrEmpty(model.Ref))
            {
                return null;
            }

            var element = new XElement(MainNs + "sortState", new XAttribute("ref", model.Ref));
            if (model.ColumnSort)
            {
                element.SetAttributeValue("columnSort", 1);
            }

            if (model.CaseSensitive)
            {
                element.SetAttributeValue("caseSensitive", 1);
            }

            if (!string.IsNullOrEmpty(model.SortMethod))
            {
                element.SetAttributeValue("sortMethod", model.SortMethod);
            }

            for (var index = 0; index < model.Conditions.Count; index++)
            {
                var condition = model.Conditions[index];
                if (string.IsNullOrEmpty(condition.Ref))
                {
                    continue;
                }

                var conditionElement = new XElement(MainNs + "sortCondition", new XAttribute("ref", condition.Ref));
                var sortBy = NormalizeSortBy(condition.SortBy);
                if (condition.Descending)
                {
                    conditionElement.SetAttributeValue("descending", 1);
                }

                if (!string.IsNullOrEmpty(sortBy))
                {
                    conditionElement.SetAttributeValue("sortBy", sortBy);
                }

                if (AllowsCustomList(sortBy) && !string.IsNullOrEmpty(condition.CustomList))
                {
                    conditionElement.SetAttributeValue("customList", condition.CustomList);
                }

                if (AllowsDifferentialStyle(sortBy) && HasValidDifferentialStyleId(condition.DifferentialStyleId, differentialFormatCount))
                {
                    conditionElement.SetAttributeValue("dxfId", condition.DifferentialStyleId.Value.ToString(CultureInfo.InvariantCulture));
                }

                if (AllowsIconSort(condition, sortBy))
                {
                    conditionElement.SetAttributeValue("iconSet", condition.IconSet);
                    conditionElement.SetAttributeValue("iconId", condition.IconId.Value.ToString(CultureInfo.InvariantCulture));
                }

                element.Add(conditionElement);
            }

            return element;
        }

        private static List<FilterColumnModel> GetOrderedFilterColumns(IReadOnlyList<FilterColumnModel> columns)
        {
            var ordered = new List<FilterColumnModel>(columns.Count);
            for (var index = 0; index < columns.Count; index++)
            {
                ordered.Add(columns[index]);
            }

            ordered.Sort(delegate(FilterColumnModel left, FilterColumnModel right)
            {
                return AutoFilterSupport.CompareFilterColumns(left, right);
            });
            return ordered;
        }

        private static FilterColumnModel LoadFilterColumn(XElement filterColumnElement, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            var columnIndex = ParseIntAttribute(filterColumnElement.Attribute("colId"));
            if (!columnIndex.HasValue || columnIndex.Value < 0)
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid filter column id.");
                }

                AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = autoFilterRange,
                });
                return null;
            }

            var model = new FilterColumnModel();
            model.ColumnIndex = columnIndex.Value;
            model.HiddenButton = ParseFilterColumnHiddenButton(filterColumnElement, diagnostics, options, sheetName, autoFilterRange);
            LoadSimpleFilters(model, filterColumnElement, diagnostics, options, sheetName, autoFilterRange);
            LoadCustomFilters(model, filterColumnElement, diagnostics, options, sheetName, autoFilterRange);
            LoadColorFilter(model.ColorFilter, filterColumnElement.Element(MainNs + "colorFilter"), stylesheet, diagnostics, options, sheetName, autoFilterRange);
            LoadDynamicFilter(model.DynamicFilter, filterColumnElement.Element(MainNs + "dynamicFilter"), diagnostics, options, sheetName, autoFilterRange);
            LoadTop10(model.Top10, filterColumnElement.Element(MainNs + "top10"), diagnostics, options, sheetName, autoFilterRange);
            return model;
        }

        private static bool ParseFilterColumnHiddenButton(XElement filterColumnElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            bool hiddenButton;
            if (TryParseBooleanAttribute(filterColumnElement.Attribute("hiddenButton"), false, out hiddenButton))
            {
                return hiddenButton;
            }

            if (filterColumnElement.Attribute("hiddenButton") != null)
            {
                return WarnAndReturnFalse(diagnostics, options, sheetName, autoFilterRange);
            }

            bool showButton;
            if (TryParseBooleanAttribute(filterColumnElement.Attribute("showButton"), true, out showButton))
            {
                return !showButton;
            }

            if (filterColumnElement.Attribute("showButton") != null)
            {
                return WarnAndReturnFalse(diagnostics, options, sheetName, autoFilterRange);
            }

            return false;
        }

        private static bool WarnAndReturnFalse(LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            if (options.StrictMode)
            {
                throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid filter column attribute.");
            }

            AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
            {
                SheetName = sheetName,
                CellRef = autoFilterRange,
            });
            return false;
        }

        private static void LoadSimpleFilters(FilterColumnModel model, XElement filterColumnElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            var filtersElement = filterColumnElement.Element(MainNs + "filters");
            if (filtersElement == null)
            {
                return;
            }

            if (filtersElement.Elements(MainNs + "dateGroupItem").Any())
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("The worksheet autoFilter contains unsupported date group filters.");
                }

                AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = autoFilterRange,
                });
            }

            foreach (var filterElement in filtersElement.Elements(MainNs + "filter"))
            {
                var value = (string)filterElement.Attribute("val");
                if (value == null)
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The worksheet autoFilter contains a filter without a value.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = autoFilterRange,
                    });
                    continue;
                }

                model.Filters.Add(value);
            }
        }

        private static void LoadCustomFilters(FilterColumnModel model, XElement filterColumnElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            var customFiltersElement = filterColumnElement.Element(MainNs + "customFilters");
            if (customFiltersElement == null)
            {
                return;
            }

            bool andValue;
            if (TryParseBooleanAttribute(customFiltersElement.Attribute("and"), false, out andValue))
            {
                model.CustomFiltersAnd = andValue;
            }
            else if (customFiltersElement.Attribute("and") != null)
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid custom filter 'and' attribute.");
                }

                AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = autoFilterRange,
                });
            }

            foreach (var customFilterElement in customFiltersElement.Elements(MainNs + "customFilter"))
            {
                var value = (string)customFilterElement.Attribute("val");
                if (value == null)
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The worksheet autoFilter contains a custom filter without a value.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = autoFilterRange,
                    });
                    continue;
                }

                FilterOperatorType operatorType;
                if (!AutoFilterSupport.TryParseOperator((string)customFilterElement.Attribute("operator"), out operatorType))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid custom filter operator.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = autoFilterRange,
                    });
                    continue;
                }

                model.CustomFilters.Add(new AutoFilterCustomFilterModel
                {
                    Operator = AutoFilterSupport.ToOperatorName(operatorType) ?? string.Empty,
                    Value = value,
                });
            }
        }

        private static void LoadColorFilter(AutoFilterColorFilterModel model, XElement colorFilterElement, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            model.Clear();
            if (colorFilterElement == null)
            {
                return;
            }

            model.Enabled = true;
            var dxfId = ParseIntAttribute(colorFilterElement.Attribute("dxfId"));
            if (colorFilterElement.Attribute("dxfId") != null)
            {
                if (HasValidDifferentialStyleId(dxfId, stylesheet.DifferentialFormats.Count))
                {
                    model.DifferentialStyleId = dxfId.Value;
                }
                else
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid color filter dxfId.");
                    }

                    AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                    {
                        SheetName = sheetName,
                        CellRef = autoFilterRange,
                    });
                    model.Clear();
                    return;
                }
            }

            bool cellColor;
            if (TryParseBooleanAttribute(colorFilterElement.Attribute("cellColor"), false, out cellColor))
            {
                model.CellColor = cellColor;
            }
            else if (colorFilterElement.Attribute("cellColor") != null)
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid color filter cellColor attribute.");
                }

                AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = autoFilterRange,
                });
            }
        }

        private static void LoadDynamicFilter(AutoFilterDynamicFilterModel model, XElement dynamicFilterElement, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            model.Clear();
            if (dynamicFilterElement == null)
            {
                return;
            }

            model.Enabled = true;
            model.Type = AutoFilterSupport.NormalizeOptionalText((string)dynamicFilterElement.Attribute("type"));
            model.Value = ParseOptionalDouble(dynamicFilterElement.Attribute("val"), diagnostics, options, sheetName, autoFilterRange, "dynamic filter value");
            model.MaxValue = ParseOptionalDouble(dynamicFilterElement.Attribute("maxVal"), diagnostics, options, sheetName, autoFilterRange, "dynamic filter max value");
        }

        private static void LoadTop10(AutoFilterTop10Model model, XElement top10Element, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            model.Clear();
            if (top10Element == null)
            {
                return;
            }

            model.Enabled = true;
            bool topValue;
            if (TryParseBooleanAttribute(top10Element.Attribute("top"), true, out topValue))
            {
                model.Top = topValue;
            }
            else if (top10Element.Attribute("top") != null)
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid top10 top attribute.");
                }

                AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = autoFilterRange,
                });
            }

            bool percentValue;
            if (TryParseBooleanAttribute(top10Element.Attribute("percent"), false, out percentValue))
            {
                model.Percent = percentValue;
            }
            else if (top10Element.Attribute("percent") != null)
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid top10 percent attribute.");
                }

                AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = autoFilterRange,
                });
            }

            model.Value = ParseOptionalDouble(top10Element.Attribute("val"), diagnostics, options, sheetName, autoFilterRange, "top10 value");
            model.FilterValue = ParseOptionalDouble(top10Element.Attribute("filterVal"), diagnostics, options, sheetName, autoFilterRange, "top10 filter value");
        }

        private static void LoadSortState(AutoFilterSortStateModel model, XElement sortStateElement, StylesheetLoadContext stylesheet, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            model.Clear();
            var reference = (string)sortStateElement.Attribute("ref");
            string normalizedRange;
            if (!AutoFilterSupport.TryNormalizeRange(reference, out normalizedRange))
            {
                if (options.StrictMode)
                {
                    throw new InvalidFileFormatException("The worksheet autoFilter sortState ref is invalid.");
                }

                AddIssue(diagnostics, options, new LoadIssue("WS-L012", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter sort state was invalid and was dropped during load.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                    CellRef = autoFilterRange,
                });
                return;
            }

            model.Ref = normalizedRange;
            bool columnSort;
            if (TryParseBooleanAttribute(sortStateElement.Attribute("columnSort"), false, out columnSort))
            {
                model.ColumnSort = columnSort;
            }
            else if (sortStateElement.Attribute("columnSort") != null)
            {
                WarnSortState(diagnostics, options, sheetName, autoFilterRange);
            }

            bool caseSensitive;
            if (TryParseBooleanAttribute(sortStateElement.Attribute("caseSensitive"), false, out caseSensitive))
            {
                model.CaseSensitive = caseSensitive;
            }
            else if (sortStateElement.Attribute("caseSensitive") != null)
            {
                WarnSortState(diagnostics, options, sheetName, autoFilterRange);
            }

            model.SortMethod = AutoFilterSupport.NormalizeOptionalText((string)sortStateElement.Attribute("sortMethod"));
            foreach (var sortConditionElement in sortStateElement.Elements(MainNs + "sortCondition"))
            {
                var conditionReference = (string)sortConditionElement.Attribute("ref");
                string normalizedConditionRef;
                if (!AutoFilterSupport.TryNormalizeRange(conditionReference, out normalizedConditionRef))
                {
                    if (options.StrictMode)
                    {
                        throw new InvalidFileFormatException("The worksheet autoFilter sortCondition ref is invalid.");
                    }

                    WarnSortState(diagnostics, options, sheetName, autoFilterRange);
                    continue;
                }

                var condition = new AutoFilterSortConditionModel();
                condition.Ref = normalizedConditionRef;
                bool descending;
                if (TryParseBooleanAttribute(sortConditionElement.Attribute("descending"), false, out descending))
                {
                    condition.Descending = descending;
                }
                else if (sortConditionElement.Attribute("descending") != null)
                {
                    WarnSortState(diagnostics, options, sheetName, autoFilterRange);
                }

                var sortBy = NormalizeSortBy((string)sortConditionElement.Attribute("sortBy"));
                if (sortConditionElement.Attribute("sortBy") != null && string.IsNullOrEmpty(sortBy))
                {
                    WarnSortState(diagnostics, options, sheetName, autoFilterRange);
                }

                condition.SortBy = sortBy;

                var customList = AutoFilterSupport.NormalizeOptionalText((string)sortConditionElement.Attribute("customList"));
                if (AllowsCustomList(sortBy))
                {
                    condition.CustomList = customList ?? string.Empty;
                }
                else if (!string.IsNullOrEmpty(customList))
                {
                    WarnSortState(diagnostics, options, sheetName, autoFilterRange);
                }

                var iconSet = AutoFilterSupport.NormalizeOptionalText((string)sortConditionElement.Attribute("iconSet"));
                if (AllowsIconSort(sortBy, iconSet))
                {
                    condition.IconSet = iconSet ?? string.Empty;
                }
                else if (!string.IsNullOrEmpty(iconSet))
                {
                    WarnSortState(diagnostics, options, sheetName, autoFilterRange);
                }

                var dxfId = ParseIntAttribute(sortConditionElement.Attribute("dxfId"));
                if (sortConditionElement.Attribute("dxfId") != null)
                {
                    if (AllowsDifferentialStyle(sortBy) && HasValidDifferentialStyleId(dxfId, stylesheet.DifferentialFormats.Count))
                    {
                        condition.DifferentialStyleId = dxfId.Value;
                    }
                    else
                    {
                        if (options.StrictMode)
                        {
                            throw new InvalidFileFormatException("The worksheet autoFilter sortCondition dxfId is invalid.");
                        }

                        WarnSortState(diagnostics, options, sheetName, autoFilterRange);
                    }
                }

                var iconId = ParseIntAttribute(sortConditionElement.Attribute("iconId"));
                if (sortConditionElement.Attribute("iconId") != null)
                {
                    if (AllowsIconSort(sortBy, iconSet) && iconId.HasValue && iconId.Value >= 0)
                    {
                        condition.IconId = iconId.Value;
                    }
                    else
                    {
                        if (options.StrictMode)
                        {
                            throw new InvalidFileFormatException("The worksheet autoFilter sortCondition iconId is invalid.");
                        }

                        WarnSortState(diagnostics, options, sheetName, autoFilterRange);
                    }
                }

                model.Conditions.Add(condition);
            }
        }

        private static void WarnSortState(LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange)
        {
            AddIssue(diagnostics, options, new LoadIssue("WS-L012", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter sort state was invalid and was dropped during load.", repairApplied: true, dataLossRisk: true)
            {
                SheetName = sheetName,
                CellRef = autoFilterRange,
            });
        }

        private static bool TryParseBooleanAttribute(XAttribute attribute, bool defaultValue, out bool parsedValue)
        {
            parsedValue = defaultValue;
            if (attribute == null)
            {
                return true;
            }

            var rawValue = ((string)attribute ?? string.Empty).Trim();
            if (rawValue == "1" || string.Equals(rawValue, "true", StringComparison.OrdinalIgnoreCase))
            {
                parsedValue = true;
                return true;
            }

            if (rawValue == "0" || string.Equals(rawValue, "false", StringComparison.OrdinalIgnoreCase))
            {
                parsedValue = false;
                return true;
            }

            return false;
        }

        private static double? ParseOptionalDouble(XAttribute attribute, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string autoFilterRange, string attributeLabel)
        {
            if (attribute == null)
            {
                return null;
            }

            double parsed;
            if (double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
            {
                return parsed;
            }

            if (options.StrictMode)
            {
                throw new InvalidFileFormatException("The worksheet autoFilter contains an invalid " + attributeLabel + ".");
            }

            AddIssue(diagnostics, options, new LoadIssue("WS-L011", DiagnosticSeverity.LossyRecoverable, "Worksheet autoFilter column metadata was invalid or conflicting and some entries were dropped during load.", repairApplied: true, dataLossRisk: true)
            {
                SheetName = sheetName,
                CellRef = autoFilterRange,
            });
            return null;
        }
        private static bool HasValidDifferentialStyleId(int? differentialStyleId, int differentialFormatCount)
        {
            return differentialStyleId.HasValue
                && differentialStyleId.Value >= 0
                && differentialStyleId.Value < differentialFormatCount;
        }
        private static string NormalizeSortBy(string sortBy)
        {
            if (string.IsNullOrWhiteSpace(sortBy))
            {
                return string.Empty;
            }

            var normalized = sortBy.Trim();
            if (string.Equals(normalized, "value", StringComparison.OrdinalIgnoreCase))
            {
                return "value";
            }

            if (string.Equals(normalized, "cellColor", StringComparison.OrdinalIgnoreCase))
            {
                return "cellColor";
            }

            if (string.Equals(normalized, "fontColor", StringComparison.OrdinalIgnoreCase))
            {
                return "fontColor";
            }

            if (string.Equals(normalized, "icon", StringComparison.OrdinalIgnoreCase))
            {
                return "icon";
            }

            return string.Empty;
        }

        private static bool AllowsCustomList(string sortBy)
        {
            return string.IsNullOrEmpty(sortBy) || string.Equals(sortBy, "value", StringComparison.Ordinal);
        }

        private static bool AllowsDifferentialStyle(string sortBy)
        {
            return string.Equals(sortBy, "cellColor", StringComparison.Ordinal)
                || string.Equals(sortBy, "fontColor", StringComparison.Ordinal);
        }

        private static bool AllowsIconSort(AutoFilterSortConditionModel condition, string sortBy)
        {
            return string.Equals(sortBy, "icon", StringComparison.Ordinal)
                && !string.IsNullOrEmpty(condition.IconSet)
                && condition.IconId.HasValue
                && condition.IconId.Value >= 0;
        }

        private static bool AllowsIconSort(string sortBy, string iconSet)
        {
            return string.Equals(sortBy, "icon", StringComparison.Ordinal)
                && !string.IsNullOrEmpty(iconSet);
        }
    }
}
