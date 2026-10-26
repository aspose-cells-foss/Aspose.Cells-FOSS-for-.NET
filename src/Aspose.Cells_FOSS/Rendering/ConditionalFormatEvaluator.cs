using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Evaluates a worksheet's conditional-formatting rules and precomputes, per cell, the visual
    /// effects the renderer should apply: a fill override, a font-color override, a data bar, or an
    /// icon-set glyph. Formula-based (<c>expression</c>) rules are skipped - there is no formula engine -
    /// and rules with no differential format contribute no visible change, matching Excel for this file.
    /// </summary>
    internal sealed class ConditionalFormatEvaluator
    {
        internal struct DataBar
        {
            public double Fraction;   // 0..1 of the cell width
            public SKColor Color;
            public bool ShowBorder;
            public bool Gradient;
        }

        internal struct Icon
        {
            public SKColor Color;
            public SKColor BorderColor;
            public bool HasBorder;
            public float AngleDeg;    // arrow direction; 90 = up, -90 = down, 0 = right
            public bool IsArrow;      // false => draw a filled circle (fallback for non-arrow sets)
        }

        internal sealed class CellEffect
        {
            public bool HasFill;
            public SKColor Fill;
            public bool HasFontColor;
            public SKColor FontColor;
            public bool HasFontBold;
            public bool HasFontItalic;
            public bool HasFontUnderline;
            public FontUnderlineType FontUnderline;
            public DataBar? Bar;
            public Icon? IconGlyph;
            public bool SuppressText;   // icon-set rule with showValue="0": draw the icon only
        }

        private static readonly SKColor ArrowGreen = new SKColor(0x63, 0xA5, 0x37);
        private static readonly SKColor ArrowAmber = new SKColor(0xE8, 0xA8, 0x38);
        private static readonly SKColor ArrowRed = new SKColor(0xDC, 0x3A, 0x2E);
        private static readonly SKColor TrafficRed = new SKColor(214, 85, 50);
        private static readonly SKColor TrafficYellow = new SKColor(234, 194, 130);
        private static readonly SKColor TrafficGreen = new SKColor(104, 164, 144);
        private static readonly SKColor TrafficRedBorder = new SKColor(161, 63, 36);
        private static readonly SKColor TrafficYellowBorder = new SKColor(167, 132, 51);
        private static readonly SKColor TrafficGreenBorder = new SKColor(58, 116, 97);

        private readonly Dictionary<long, CellEffect> _effects = new Dictionary<long, CellEffect>();

        public ConditionalFormatEvaluator(WorksheetModel sheet, RenderContext context)
        {
            if (sheet == null || sheet.ConditionalFormattings == null)
            {
                return;
            }

            // Higher priority (lower number) wins for overlapping differential-format rules.
            var formattings = new List<ConditionalFormattingModel>(sheet.ConditionalFormattings);
            foreach (var formatting in formattings)
            {
                var conditions = new List<FormatConditionModel>(formatting.Conditions);
                conditions.Sort((a, b) => a.Priority.CompareTo(b.Priority));

                foreach (var condition in conditions)
                {
                    foreach (var area in formatting.Areas)
                    {
                        ApplyConditionToArea(sheet, context, condition, area);
                    }
                }
            }
        }

        public bool TryGet(int row, int col, out CellEffect effect)
        {
            return _effects.TryGetValue(Key(row, col), out effect);
        }

        private CellEffect EffectFor(int row, int col)
        {
            CellEffect effect;
            var key = Key(row, col);
            if (!_effects.TryGetValue(key, out effect))
            {
                effect = new CellEffect();
                _effects[key] = effect;
            }

            return effect;
        }

        private void ApplyConditionToArea(WorksheetModel sheet, RenderContext context, FormatConditionModel condition, CellArea area)
        {
            switch (condition.Type)
            {
                case FormatConditionType.ColorScale:
                    ApplyColorScale(sheet, context, condition, area);
                    return;
                case FormatConditionType.DataBar:
                    ApplyDataBar(sheet, context, condition, area);
                    return;
                case FormatConditionType.IconSet:
                    ApplyIconSet(sheet, condition, area);
                    return;
                case FormatConditionType.Expression:
                    ApplyExpression(sheet, context, condition, area);
                    return;
            }

            // Differential-format (dxf) rules: apply the fill/font only where the rule matches. A rule
            // with no dxf still "matches" but paints nothing (Excel behaviour for this file).
            var matcher = BuildMatcher(sheet, context, condition, area);
            if (matcher == null)
            {
                return;
            }

            for (var r = area.StartRow; r <= area.EndRow; r++)
            {
                for (var c = area.StartColumn; c <= area.EndColumn; c++)
                {
                    CellRecord record;
                    sheet.Cells.TryGetValue(new CellAddress(r, c), out record);
                    if (!matcher(record))
                    {
                        continue;
                    }

                    ApplyDifferentialStyle(context, EffectFor(r, c), condition);
                }
            }
        }

        private void ApplyExpression(WorksheetModel sheet, RenderContext context,
            FormatConditionModel condition, CellArea area)
        {
            var matcher = ConditionalFormatExpressionMatcher.TryCreate(sheet, area, condition.Formula1);
            if (matcher == null)
            {
                return;
            }

            for (var row = area.StartRow; row <= area.EndRow; row++)
            {
                for (var column = area.StartColumn; column <= area.EndColumn; column++)
                {
                    if (matcher.Matches(row, column))
                    {
                        ApplyDifferentialStyle(context, EffectFor(row, column), condition);
                    }
                }
            }
        }

        private static void ApplyDifferentialStyle(RenderContext context, CellEffect effect, FormatConditionModel condition)
        {
            var style = condition.Style;
            if (style == null)
            {
                return;
            }

            if (!effect.HasFill && style.Pattern != FillPatternKind.None)
            {
                var fill = context.Colors.Resolve(style.ForegroundColor, SKColors.Transparent);
                if (fill.Alpha != 0)
                {
                    effect.HasFill = true;
                    effect.Fill = fill;
                }
            }

            if (!effect.HasFontColor && style.Font != null && style.Font.Color.A != 0)
            {
                effect.HasFontColor = true;
                effect.FontColor = context.Colors.Resolve(style.Font.Color, SKColors.Black);
            }

            if (style.Font != null && style.Font.Bold)
            {
                effect.HasFontBold = true;
            }

            if (style.Font != null && style.Font.Italic)
            {
                effect.HasFontItalic = true;
            }

            if (style.Font != null && style.Font.Underline != FontUnderlineType.None)
            {
                effect.HasFontUnderline = true;
                effect.FontUnderline = style.Font.Underline;
            }
        }

        // --- Differential-format rule matchers -----------------------------------------------------

        private delegate bool CellMatcher(CellRecord record);

        private static CellMatcher BuildMatcher(WorksheetModel sheet, RenderContext context,
            FormatConditionModel condition, CellArea area)
        {
            switch (condition.Type)
            {
                case FormatConditionType.ContainsText:
                {
                    var needle = condition.Formula1 ?? string.Empty;
                    return record => TextOf(record).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
                }
                case FormatConditionType.NotContainsText:
                {
                    var needle = condition.Formula1 ?? string.Empty;
                    return record => TextOf(record).IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0;
                }
                case FormatConditionType.BeginsWith:
                {
                    var prefix = condition.Formula1 ?? string.Empty;
                    return record => TextOf(record).StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
                }
                case FormatConditionType.EndsWith:
                {
                    var suffix = condition.Formula1 ?? string.Empty;
                    return record => TextOf(record).EndsWith(suffix, StringComparison.OrdinalIgnoreCase);
                }
                case FormatConditionType.TimePeriod:
                    return BuildTimePeriodMatcher(condition.TimePeriod, context);
                case FormatConditionType.CellValue:
                    return BuildCellValueMatcher(condition);
                case FormatConditionType.DuplicateValues:
                case FormatConditionType.UniqueValues:
                    return BuildUniquenessMatcher(sheet, area, condition.Type == FormatConditionType.DuplicateValues);
                case FormatConditionType.Top10:
                case FormatConditionType.Bottom10:
                    return BuildTopMatcher(sheet, area, condition);
                case FormatConditionType.AboveAverage:
                case FormatConditionType.BelowAverage:
                    return BuildAverageMatcher(sheet, area, condition);
                default:
                    // Unmodelled rule types have no evaluator.
                    return null;
            }
        }

        private static CellMatcher BuildCellValueMatcher(FormatConditionModel condition)
        {
            double v1, v2;
            var has1 = TryParseDouble(condition.Formula1, out v1);
            var has2 = TryParseDouble(condition.Formula2, out v2);
            var op = condition.Operator;
            return record =>
            {
                double x;
                if (!TryNumber(record, out x))
                {
                    return false;
                }

                switch (op)
                {
                    case OperatorType.GreaterThan: return has1 && x > v1;
                    case OperatorType.GreaterOrEqual: return has1 && x >= v1;
                    case OperatorType.LessThan: return has1 && x < v1;
                    case OperatorType.LessOrEqual: return has1 && x <= v1;
                    case OperatorType.Equal: return has1 && x == v1;
                    case OperatorType.NotEqual: return has1 && x != v1;
                    case OperatorType.Between: return has1 && has2 && x >= Math.Min(v1, v2) && x <= Math.Max(v1, v2);
                    case OperatorType.NotBetween: return has1 && has2 && (x < Math.Min(v1, v2) || x > Math.Max(v1, v2));
                    default: return false;
                }
            };
        }

        private static CellMatcher BuildTimePeriodMatcher(string period, RenderContext context)
        {
            var today = context.ConditionalFormattingEvaluationDate.Date;
            var dateSystem = context.Workbook.Settings.DateSystem;
            return record =>
            {
                double serial;
                if (!TryNumber(record, out serial))
                {
                    return false;
                }

                var date = SerialToDate(serial, dateSystem);
                if (!date.HasValue)
                {
                    return false;
                }

                var d = date.Value.Date;
                switch ((period ?? string.Empty).ToLowerInvariant())
                {
                    case "today": return d == today;
                    case "yesterday": return d == today.AddDays(-1);
                    case "tomorrow": return d == today.AddDays(1);
                    case "last7days": return d <= today && d >= today.AddDays(-6);
                    case "thismonth": return d.Year == today.Year && d.Month == today.Month;
                    case "lastmonth": { var m = today.AddMonths(-1); return d.Year == m.Year && d.Month == m.Month; }
                    case "nextmonth": { var m = today.AddMonths(1); return d.Year == m.Year && d.Month == m.Month; }
                    default: return false;
                }
            };
        }

        private static CellMatcher BuildUniquenessMatcher(WorksheetModel sheet, CellArea area, bool wantDuplicates)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            ForEachValue(sheet, area, (record, text) =>
            {
                int n;
                counts.TryGetValue(text, out n);
                counts[text] = n + 1;
            });

            return record =>
            {
                var text = ValueKey(record);
                if (text == null)
                {
                    return false;
                }

                int n;
                counts.TryGetValue(text, out n);
                return wantDuplicates ? n > 1 : n == 1;
            };
        }

        private static CellMatcher BuildTopMatcher(WorksheetModel sheet, CellArea area, FormatConditionModel condition)
        {
            var values = CollectNumbers(sheet, area);
            if (values.Count == 0 || condition.Rank <= 0)
            {
                return record => false;
            }

            var count = condition.Percent
                ? (int)Math.Ceiling(values.Count * (condition.Rank / 100.0))
                : condition.Rank;
            count = Math.Max(1, Math.Min(count, values.Count));

            values.Sort();
            double threshold = condition.Top
                ? values[values.Count - count]   // keep the largest `count`
                : values[count - 1];             // keep the smallest `count`

            return record =>
            {
                double x;
                if (!TryNumber(record, out x))
                {
                    return false;
                }

                return condition.Top ? x >= threshold : x <= threshold;
            };
        }

        private static CellMatcher BuildAverageMatcher(WorksheetModel sheet, CellArea area, FormatConditionModel condition)
        {
            var values = CollectNumbers(sheet, area);
            if (values.Count == 0)
            {
                return record => false;
            }

            double sum = 0;
            foreach (var v in values) sum += v;
            var mean = sum / values.Count;

            var cutoff = mean;
            if (condition.StandardDeviation > 0)
            {
                double sq = 0;
                foreach (var v in values) sq += (v - mean) * (v - mean);
                var sd = Math.Sqrt(sq / values.Count);
                cutoff = condition.Above ? mean + condition.StandardDeviation * sd : mean - condition.StandardDeviation * sd;
            }

            return record =>
            {
                double x;
                if (!TryNumber(record, out x))
                {
                    return false;
                }

                return condition.Above ? x > cutoff : x < cutoff;
            };
        }

        // --- Color scale / data bar / icon set -----------------------------------------------------

        private void ApplyColorScale(WorksheetModel sheet, RenderContext context, FormatConditionModel condition, CellArea area)
        {
            var values = CollectNumbers(sheet, area);
            if (values.Count == 0)
            {
                return;
            }

            values.Sort();
            var min = values[0];
            var max = values[values.Count - 1];
            var threeColor = condition.ColorScaleCount >= 3 && condition.MidColor.A != 0;
            var mid = threeColor ? Percentile(values, 50) : (min + max) / 2.0;

            var minColor = ToSk(context, condition.MinColor, new SKColor(0xF8, 0x69, 0x6B));
            var midColor = ToSk(context, condition.MidColor, new SKColor(0xFF, 0xEB, 0x84));
            var maxColor = ToSk(context, condition.MaxColor, new SKColor(0x63, 0xBE, 0x7B));

            for (var r = area.StartRow; r <= area.EndRow; r++)
            {
                for (var c = area.StartColumn; c <= area.EndColumn; c++)
                {
                    CellRecord record;
                    sheet.Cells.TryGetValue(new CellAddress(r, c), out record);
                    double x;
                    if (!TryNumber(record, out x))
                    {
                        continue;
                    }

                    SKColor color;
                    if (!threeColor)
                    {
                        color = Lerp(minColor, maxColor, Fraction(x, min, max));
                    }
                    else if (x <= mid)
                    {
                        color = Lerp(minColor, midColor, Fraction(x, min, mid));
                    }
                    else
                    {
                        color = Lerp(midColor, maxColor, Fraction(x, mid, max));
                    }

                    var effect = EffectFor(r, c);
                    effect.HasFill = true;
                    effect.Fill = color;
                }
            }
        }

        private void ApplyDataBar(WorksheetModel sheet, RenderContext context, FormatConditionModel condition, CellArea area)
        {
            var values = CollectNumbers(sheet, area);
            if (values.Count == 0)
            {
                return;
            }

            values.Sort();
            var min = values[0];
            var max = values[values.Count - 1];
            if (condition.DataBarAutoMinimum && min >= 0d)
            {
                min = 0d;
            }
            var color = ToSk(context, condition.BarColor, new SKColor(0x63, 0x8E, 0xC6));

            for (var r = area.StartRow; r <= area.EndRow; r++)
            {
                for (var c = area.StartColumn; c <= area.EndColumn; c++)
                {
                    CellRecord record;
                    sheet.Cells.TryGetValue(new CellAddress(r, c), out record);
                    double x;
                    if (!TryNumber(record, out x))
                    {
                        continue;
                    }

                    var display = condition.DataBarMinLength
                        + (condition.DataBarMaxLength - condition.DataBarMinLength) * Fraction(x, min, max);
                    var effect = EffectFor(r, c);
                    effect.Bar = new DataBar { Fraction = display, Color = color,
                        ShowBorder = condition.ShowBorder, Gradient = condition.DataBarGradient };
                }
            }
        }

        private void ApplyIconSet(WorksheetModel sheet, FormatConditionModel condition, CellArea area)
        {
            var values = CollectNumbers(sheet, area);
            if (values.Count == 0)
            {
                return;
            }

            values.Sort();
            var min = values[0];
            var max = values[values.Count - 1];
            var iconCount = IconCount(condition.IconSetType);

            for (var r = area.StartRow; r <= area.EndRow; r++)
            {
                for (var c = area.StartColumn; c <= area.EndColumn; c++)
                {
                    CellRecord record;
                    sheet.Cells.TryGetValue(new CellAddress(r, c), out record);
                    double x;
                    if (!TryNumber(record, out x))
                    {
                        continue;
                    }

                    // OOXML's default percent thresholds are integer values (0, 33, 66 for three icons).
                    var percent = Fraction(x, min, max) * 100.0;
                    var bucket = 0;
                    for (var index = 1; index < iconCount; index++)
                    {
                        if (percent >= (100 * index) / iconCount)
                        {
                            bucket = index;
                        }
                    }

                    var lookup = condition.ReverseIcons ? iconCount - 1 - bucket : bucket;
                    var effect = EffectFor(r, c);
                    effect.IconGlyph = IconFor(condition.IconSetType, lookup, iconCount);
                    if (condition.ShowIconOnly)
                    {
                        effect.SuppressText = true;
                    }
                }
            }
        }

        private static Icon IconFor(string iconSetType, int index, int count)
        {
            var type = (iconSetType ?? string.Empty).ToLowerInvariant();
            var isArrow = type.Contains("arrow");
            var isTrafficLight = type.Contains("trafficlight");

            // Colour ramp low->high: red, amber, green (amber fills the middle buckets).
            SKColor color;
            if (index == 0) color = ArrowRed;
            else if (index == count - 1) color = ArrowGreen;
            else color = ArrowAmber;

            var borderColor = SKColors.Transparent;
            if (isTrafficLight)
            {
                if (index == 0)
                {
                    color = TrafficRed;
                    borderColor = TrafficRedBorder;
                }
                else if (index == count - 1)
                {
                    color = TrafficGreen;
                    borderColor = TrafficGreenBorder;
                }
                else
                {
                    color = TrafficYellow;
                    borderColor = TrafficYellowBorder;
                }
            }

            // Arrow direction sweeps from down (lowest) to up (highest) across the buckets.
            var angle = count <= 1 ? 0f : -90f + 180f * index / (count - 1);
            return new Icon { Color = color, BorderColor = borderColor,
                HasBorder = isTrafficLight, AngleDeg = angle, IsArrow = isArrow };
        }

        // --- Value / numeric helpers ---------------------------------------------------------------

        private static void ForEachValue(WorksheetModel sheet, CellArea area, Action<CellRecord, string> action)
        {
            for (var r = area.StartRow; r <= area.EndRow; r++)
            {
                for (var c = area.StartColumn; c <= area.EndColumn; c++)
                {
                    CellRecord record;
                    if (!sheet.Cells.TryGetValue(new CellAddress(r, c), out record))
                    {
                        continue;
                    }

                    var key = ValueKey(record);
                    if (key != null)
                    {
                        action(record, key);
                    }
                }
            }
        }

        private static List<double> CollectNumbers(WorksheetModel sheet, CellArea area)
        {
            var list = new List<double>();
            for (var r = area.StartRow; r <= area.EndRow; r++)
            {
                for (var c = area.StartColumn; c <= area.EndColumn; c++)
                {
                    CellRecord record;
                    sheet.Cells.TryGetValue(new CellAddress(r, c), out record);
                    double x;
                    if (TryNumber(record, out x))
                    {
                        list.Add(x);
                    }
                }
            }

            return list;
        }

        private static string ValueKey(CellRecord record)
        {
            if (record == null || record.Value == null)
            {
                return null;
            }

            double x;
            if (TryNumber(record, out x))
            {
                return "n:" + x.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            var text = TextOf(record);
            return text.Length == 0 ? null : "s:" + text;
        }

        private static string TextOf(CellRecord record)
        {
            if (record == null || record.Value == null)
            {
                return string.Empty;
            }

            return Convert.ToString(record.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
        }

        internal static bool TryNumber(CellRecord record, out double value)
        {
            value = 0;
            if (record == null || record.Value == null)
            {
                return false;
            }

            if (record.Kind == CellValueKind.String || record.Kind == CellValueKind.Boolean || record.Kind == CellValueKind.Error)
            {
                return false;
            }

            return TryToDouble(record.Value, out value);
        }

        private static bool TryToDouble(object value, out double result)
        {
            result = 0;
            if (value is double) { result = (double)value; return true; }
            if (value is float) { result = (float)value; return true; }
            if (value is decimal) { result = (double)(decimal)value; return true; }
            if (value is int) { result = (int)value; return true; }
            if (value is long) { result = (long)value; return true; }
            if (value is short) { result = (short)value; return true; }
            if (value is byte) { result = (byte)value; return true; }
            if (value is DateTime) { result = ((DateTime)value).ToOADate(); return true; }
            return false;
        }

        private static bool TryParseDouble(string text, out double value)
        {
            return double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value);
        }

        private static DateTime? SerialToDate(double serial, DateSystem dateSystem)
        {
            try
            {
                return DateSerialConverter.FromSerial(serial, dateSystem);
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
            catch (OverflowException)
            {
                return null;
            }
        }

        private static double Fraction(double value, double min, double max)
        {
            if (max <= min)
            {
                return 0;
            }

            var f = (value - min) / (max - min);
            if (f < 0) return 0;
            if (f > 1) return 1;
            return f;
        }

        private static double Percentile(List<double> sorted, double percent)
        {
            if (sorted.Count == 1)
            {
                return sorted[0];
            }

            var rank = percent / 100.0 * (sorted.Count - 1);
            var lo = (int)Math.Floor(rank);
            var hi = (int)Math.Ceiling(rank);
            if (lo == hi)
            {
                return sorted[lo];
            }

            return sorted[lo] + (rank - lo) * (sorted[hi] - sorted[lo]);
        }

        private static SKColor ToSk(RenderContext context, ColorValue color, SKColor fallback)
        {
            if (color.A == 0 && !color.ThemeIndex.HasValue)
            {
                return fallback;
            }

            return context.Colors.Resolve(color, fallback);
        }

        private static SKColor Lerp(SKColor a, SKColor b, double t)
        {
            return new SKColor(
                (byte)(a.Red + (int)((b.Red - a.Red) * t)),
                (byte)(a.Green + (int)((b.Green - a.Green) * t)),
                (byte)(a.Blue + (int)((b.Blue - a.Blue) * t)));
        }

        private static int IconCount(string iconSetType)
        {
            if (string.IsNullOrEmpty(iconSetType))
            {
                return 3;
            }

            if (iconSetType[0] == '4') return 4;
            if (iconSetType[0] == '5') return 5;
            return 3;
        }

        private static long Key(int row, int col)
        {
            return ((long)row << 20) | (uint)col;
        }
    }
}
