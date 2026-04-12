using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;
using static Aspose.Cells_FOSS.XlsxWorkbookStylesXml;

namespace Aspose.Cells_FOSS;

internal static class XlsxWorkbookStyles
{

    internal static StylesheetSaveContext BuildStylesheet(WorkbookModel model)
    {
        var defaultStyle = model.DefaultStyle.Clone();

        var fonts = new List<FontValue> { defaultStyle.Font.Clone() };
        var fontIds = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            [GetFontKey(fonts[0])] = 0,
        };

        var fills = new List<FillValue>
        {
            ToFillValue(defaultStyle),
            new FillValue { Pattern = FillPatternKind.Gray125 },
        };
        var fillIds = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            [GetFillKey(fills[0])] = 0,
            [GetFillKey(fills[1])] = 1,
        };

        var borders = new List<BordersValue> { defaultStyle.Borders.Clone() };
        var borderIds = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            [GetBordersKey(borders[0])] = 0,
        };

        var customNumberFormatIds = new Dictionary<string, int>(StringComparer.Ordinal);
        var customNumberFormats = new List<KeyValuePair<int, string>>();
        var nextCustomNumberFormatId = 164;
        var defaultNumberFormatId = ResolveNumberFormatId(defaultStyle.NumberFormat, customNumberFormatIds, customNumberFormats, ref nextCustomNumberFormatId);

        var normalCellFormat = new CellFormatValue
        {
            NumFmtId = defaultNumberFormatId,
            FontId = 0,
            FillId = 0,
            BorderId = 0,
            Alignment = defaultStyle.Alignment.Clone(),
            Protection = defaultStyle.Protection.Clone(),
        };

        var cellFormats = new List<CellFormatValue>
        {
            normalCellFormat,
        };
        var styleIndices = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            [GetStyleKey(defaultStyle)] = 0,
        };
        var differentialFormats = new List<StyleValue>();
        var differentialStyleIndices = new Dictionary<string, int>(StringComparer.Ordinal);

        foreach (var worksheet in model.Worksheets)
        {
            var persistedEntries = CollectPersistedCellEntries(worksheet, model.DefaultStyle);
            for (var index = 0; index < persistedEntries.Count; index++)
            {
                var pair = persistedEntries[index];
                var style = GetStyleForSerialization(pair.Value);
                var styleKey = GetStyleKey(style);
                if (styleIndices.ContainsKey(styleKey))
                {
                    continue;
                }

                var fontId = Intern(fontIds, fonts, style.Font, GetFontKey);
                var fillId = Intern(fillIds, fills, ToFillValue(style), GetFillKey);
                var borderId = Intern(borderIds, borders, style.Borders, GetBordersKey);
                var numFmtId = ResolveNumberFormatId(style.NumberFormat, customNumberFormatIds, customNumberFormats, ref nextCustomNumberFormatId);

                styleIndices[styleKey] = cellFormats.Count;
                cellFormats.Add(new CellFormatValue
                {
                    NumFmtId = numFmtId,
                    FontId = fontId,
                    FillId = fillId,
                    BorderId = borderId,
                    Alignment = style.Alignment.Clone(),
                    Protection = style.Protection.Clone(),
                });
            }

            for (var formattingIndex = 0; formattingIndex < worksheet.ConditionalFormattings.Count; formattingIndex++)
            {
                var formatting = worksheet.ConditionalFormattings[formattingIndex];
                for (var conditionIndex = 0; conditionIndex < formatting.Conditions.Count; conditionIndex++)
                {
                    var style = formatting.Conditions[conditionIndex].Style;
                    if (StylesEqual(style, StyleValue.Default))
                    {
                        continue;
                    }

                    var styleKey = GetStyleKey(style);
                    if (differentialStyleIndices.ContainsKey(styleKey))
                    {
                        continue;
                    }

                    differentialStyleIndices[styleKey] = differentialFormats.Count;
                    differentialFormats.Add(style.Clone());
                }
            }
        }

        var hasStyles = styleIndices.Count > 1 || !StylesEqual(defaultStyle, StyleValue.Default) || differentialFormats.Count > 0;
        return new StylesheetSaveContext(BuildStylesheetDocument(fonts, fills, borders, normalCellFormat, cellFormats, customNumberFormats, differentialFormats), styleIndices, differentialStyleIndices, differentialFormats.Count, hasStyles);
    }

    internal static StyleValue GetStyleForSerialization(CellRecord record)
    {
        var style = GetEffectiveStyle(record);
        if (record.Kind != CellValueKind.DateTime || IsDateNumberFormat(style.NumberFormat))
        {
            return style;
        }

        var serializedStyle = style.Clone();
        serializedStyle.NumberFormat = new NumberFormatValue
        {
            Number = 14,
            Custom = null,
        };
        return serializedStyle;
    }

    private static List<KeyValuePair<CellAddress, CellRecord>> CollectPersistedCellEntries(WorksheetModel worksheet, StyleValue workbookDefaultStyle)
    {
        var persistedEntries = new List<KeyValuePair<CellAddress, CellRecord>>();
        foreach (var pair in worksheet.Cells)
        {
            if (ShouldPersistCell(workbookDefaultStyle, pair.Value))
            {
                persistedEntries.Add(pair);
            }
        }

        persistedEntries.Sort(CompareCellEntries);
        return persistedEntries;
    }

    private static int CompareCellEntries(KeyValuePair<CellAddress, CellRecord> left, KeyValuePair<CellAddress, CellRecord> right)
    {
        var rowComparison = left.Key.RowIndex.CompareTo(right.Key.RowIndex);
        if (rowComparison != 0)
        {
            return rowComparison;
        }

        return left.Key.ColumnIndex.CompareTo(right.Key.ColumnIndex);
    }

    internal static StylesheetLoadContext LoadStylesheet(ZipArchive archive, IReadOnlyDictionary<string, string> workbookRelationships, LoadOptions options, LoadDiagnostics diagnostics)
    {
        var stylesUri = FindRelationshipTarget(workbookRelationships, "/xl/styles.xml");
        if (string.IsNullOrEmpty(stylesUri))
        {
            if (workbookRelationships.Count == 0)
            {
                stylesUri = "/xl/styles.xml";
            }
            else
            {
                stylesUri = FindRelationshipTargetByType(workbookRelationships, StylesRelationshipType);
            }
        }

        if (string.IsNullOrEmpty(stylesUri))
        {
            stylesUri = "/xl/styles.xml";
        }

        var entry = GetEntry(archive, stylesUri);
        if (entry is null)
        {
            return new StylesheetLoadContext();
        }

        var document = LoadDocument(entry);
        var root = document.Root;
        if (root is null)
        {
            return new StylesheetLoadContext();
        }

        var customFormats = new Dictionary<int, string>();
        foreach (var numFmt in root.Element(MainNs + "numFmts")?.Elements(MainNs + "numFmt") ?? Enumerable.Empty<XElement>())
        {
            var id = ParseIntAttribute(numFmt.Attribute("numFmtId"));
            var code = (string?)numFmt.Attribute("formatCode");
            if (id.HasValue && !string.IsNullOrEmpty(code))
            {
                customFormats[id.Value] = code!;
            }
        }

        var fonts = ReadFontValues(root);
        if (fonts.Count == 0)
        {
            fonts.Add(StyleValue.Default.Font.Clone());
        }

        var fills = ReadFillValues(root);
        if (fills.Count == 0)
        {
            fills.Add(new FillValue { Pattern = FillPatternKind.None });
            fills.Add(new FillValue { Pattern = FillPatternKind.Gray125 });
        }

        var borders = ReadBordersValues(root);
        if (borders.Count == 0)
        {
            borders.Add(StyleValue.Default.Borders.Clone());
        }

        var context = new StylesheetLoadContext();
        context.CellFormats.Clear();
        context.DifferentialFormats.AddRange(ReadDifferentialStyleValues(root));

        var cellXfs = root.Element(MainNs + "cellXfs")?.Elements(MainNs + "xf").ToList() ?? new List<XElement>();
        if (cellXfs.Count == 0)
        {
            context.CellFormats.Add(StyleValue.Default.Clone());
            context.DefaultCellStyle = context.CellFormats[0].Clone();
            return context;
        }

        for (var index = 0; index < cellXfs.Count; index++)
        {
            var xf = cellXfs[index];
            var style = StyleValue.Default.Clone();

            var fontId = ParseIntAttribute(xf.Attribute("fontId"));
            if (fontId.HasValue && fontId.Value >= 0 && fontId.Value < fonts.Count)
            {
                style.Font = fonts[fontId.Value].Clone();
            }

            var fillId = ParseIntAttribute(xf.Attribute("fillId"));
            if (fillId.HasValue && fillId.Value >= 0 && fillId.Value < fills.Count)
            {
                var fill = fills[fillId.Value];
                style.Pattern = fill.Pattern;
                style.ForegroundColor = fill.ForegroundColor;
                style.BackgroundColor = fill.BackgroundColor;
            }

            var borderId = ParseIntAttribute(xf.Attribute("borderId"));
            if (borderId.HasValue && borderId.Value >= 0 && borderId.Value < borders.Count)
            {
                style.Borders = borders[borderId.Value].Clone();
            }

            var numFmtId = ParseIntAttribute(xf.Attribute("numFmtId")) ?? 0;
            style.NumberFormat = new NumberFormatValue
            {
                Number = numFmtId,
                Custom = customFormats.TryGetValue(numFmtId, out var customFormat) ? customFormat : null,
            };

            var alignment = xf.Element(MainNs + "alignment");
            if (alignment is not null)
            {
                style.Alignment = new AlignmentValue
                {
                    Horizontal = ParseHorizontalAlignment((string?)alignment.Attribute("horizontal")),
                    Vertical = ParseVerticalAlignment((string?)alignment.Attribute("vertical")),
                    WrapText = ParseBoolAttribute(alignment.Attribute("wrapText")),
                    IndentLevel = StyleValueSanitizer.NormalizeIndentLevel(ParseIntAttribute(alignment.Attribute("indent"))),
                    TextRotation = StyleValueSanitizer.NormalizeTextRotation(ParseIntAttribute(alignment.Attribute("textRotation"))),
                    ShrinkToFit = ParseBoolAttribute(alignment.Attribute("shrinkToFit")),
                    ReadingOrder = StyleValueSanitizer.NormalizeReadingOrder(ParseIntAttribute(alignment.Attribute("readingOrder"))),
                    RelativeIndent = ParseIntAttribute(alignment.Attribute("relativeIndent")) ?? 0,
                };
            }

            var protection = xf.Element(MainNs + "protection");
            if (protection is not null)
            {
                style.Protection = new ProtectionValue
                {
                    IsLocked = ParseOptionalBoolAttribute(protection.Attribute("locked")) ?? true,
                    IsHidden = ParseBoolAttribute(protection.Attribute("hidden")),
                };
            }

            context.CellFormats.Add(style);
            if (IsDateNumberFormat(style.NumberFormat))
            {
                context.DateStyleIndexes.Add(index);
            }
        }

        if (context.CellFormats.Count == 0)
        {
            context.CellFormats.Add(StyleValue.Default.Clone());
        }

        context.DefaultCellStyle = context.CellFormats[0].Clone();
        return context;
    }

    private static int Intern<T>(IDictionary<string, int> indices, ICollection<T> values, T value, Func<T, string> getKey)
    {
        var key = getKey(value);
        if (indices.TryGetValue(key, out var existing))
        {
            return existing;
        }

        var id = values.Count;
        values.Add(value);
        indices[key] = id;
        return id;
    }

    internal static string GetStyleKey(StyleValue style)
    {
        return string.Concat(
            GetFontKey(style.Font), '|',
            GetFillKey(ToFillValue(style)), '|',
            GetBordersKey(style.Borders), '|',
            style.NumberFormat.Number, '|', style.NumberFormat.Custom ?? string.Empty, '|',
            (int)style.Alignment.Horizontal, '|', (int)style.Alignment.Vertical, '|', style.Alignment.WrapText ? '1' : '0', '|',
            style.Alignment.IndentLevel.ToString(CultureInfo.InvariantCulture), '|',
            style.Alignment.TextRotation.ToString(CultureInfo.InvariantCulture), '|',
            style.Alignment.ShrinkToFit ? '1' : '0', '|',
            style.Alignment.ReadingOrder.ToString(CultureInfo.InvariantCulture), '|',
            style.Alignment.RelativeIndent.ToString(CultureInfo.InvariantCulture), '|',
            style.Protection.IsLocked ? '1' : '0', '|', style.Protection.IsHidden ? '1' : '0');
    }

    private static string GetFontKey(FontValue font)
    {
        return string.Concat(
            font.Name, '|', font.Size.ToString("0.####", CultureInfo.InvariantCulture), '|',
            font.Bold ? '1' : '0', '|', font.Italic ? '1' : '0', '|', font.Underline ? '1' : '0', '|', font.StrikeThrough ? '1' : '0', '|',
            GetColorKey(font.Color));
    }

    private static string GetFillKey(FillValue fill)
    {
        return string.Concat((int)fill.Pattern, '|', GetColorKey(fill.ForegroundColor), '|', GetColorKey(fill.BackgroundColor));
    }

    private static string GetBordersKey(BordersValue borders)
    {
        return string.Concat(
            GetBorderSideKey(borders.Left), '|',
            GetBorderSideKey(borders.Right), '|',
            GetBorderSideKey(borders.Top), '|',
            GetBorderSideKey(borders.Bottom), '|',
            GetBorderSideKey(borders.Diagonal), '|',
            borders.DiagonalUp ? '1' : '0', '|',
            borders.DiagonalDown ? '1' : '0');
    }

    private static string GetBorderSideKey(BorderSideValue border)
    {
        return string.Concat((int)border.Style, '|', GetColorKey(border.Color));
    }

    private static string GetColorKey(ColorValue color)
    {
        return ToArgbHex(color);
    }

    private static string ToArgbHex(ColorValue color)
    {
        return string.Concat(
            color.A.ToString("X2", CultureInfo.InvariantCulture),
            color.R.ToString("X2", CultureInfo.InvariantCulture),
            color.G.ToString("X2", CultureInfo.InvariantCulture),
            color.B.ToString("X2", CultureInfo.InvariantCulture));
    }

    private static StyleValue GetEffectiveStyle(CellRecord record)
    {
        if (record.Style is null)
        {
            return StyleValue.Default.Clone();
        }

        return record.Style;
    }

    private static FillValue ToFillValue(StyleValue style)
    {
        return new FillValue
        {
            Pattern = style.Pattern,
            ForegroundColor = style.ForegroundColor,
            BackgroundColor = style.BackgroundColor,
        };
    }

    private static int ResolveNumberFormatId(
        NumberFormatValue numberFormat,
        IDictionary<string, int> customNumberFormatIds,
        ICollection<KeyValuePair<int, string>> customNumberFormats,
        ref int nextCustomNumberFormatId)
    {
        var customFormat = numberFormat.Custom;
        if (!string.IsNullOrEmpty(customFormat))
        {
            if (customNumberFormatIds.TryGetValue(customFormat!, out var existingId))
            {
                return existingId;
            }

            var customId = nextCustomNumberFormatId++;
            customNumberFormatIds[customFormat!] = customId;
            customNumberFormats.Add(new KeyValuePair<int, string>(customId, customFormat!));
            return customId;
        }

        return numberFormat.Number >= 0 ? numberFormat.Number : 0;
    }

    internal static bool StylesEqual(StyleValue left, StyleValue right)
    {
        return string.Equals(GetStyleKey(left), GetStyleKey(right), StringComparison.Ordinal);
    }

    private static bool IsDateNumberFormat(NumberFormatValue numberFormat)
    {
        if (BuiltInDateFormats.Contains(numberFormat.Number))
        {
            return true;
        }

        var customFormat = numberFormat.Custom;
        return !string.IsNullOrEmpty(customFormat) && LooksLikeDateFormat(customFormat!);
    }
}
