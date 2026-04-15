using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.Text;

namespace Aspose.Cells_FOSS
{
    internal static class DisplayTextDateFormatSupport
    {
        internal static string FormatDateTimeValue(DateTime value, string formatCode, CultureInfo culture)
        {
            var builder = new StringBuilder(formatCode.Length + 16);
            var hasAmPm = formatCode.IndexOf("AM/PM", StringComparison.OrdinalIgnoreCase) >= 0
                || formatCode.IndexOf("A/P", StringComparison.OrdinalIgnoreCase) >= 0;
            var inQuote = false;

            for (var index = 0; index < formatCode.Length; index++)
            {
                if (MatchesToken(formatCode, index, "AM/PM"))
                {
                    builder.Append(GetAmPmDesignator(value, culture, false));
                    index += 4;
                    continue;
                }

                if (MatchesToken(formatCode, index, "A/P"))
                {
                    builder.Append(GetAmPmDesignator(value, culture, true));
                    index += 2;
                    continue;
                }

                var character = formatCode[index];
                if (character == '"')
                {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote)
                {
                    builder.Append(character);
                    continue;
                }

                if (character == '\\')
                {
                    if (index + 1 < formatCode.Length)
                    {
                        index++;
                        builder.Append(formatCode[index]);
                    }

                    continue;
                }

                if (character == '_' || character == '*')
                {
                    index++;
                    continue;
                }

                if (character == '.' && TryAppendFractionalSeconds(value, formatCode, culture, ref index, builder))
                {
                    continue;
                }

                if (character == 'y' || character == 'Y')
                {
                    var count = CountRepeated(formatCode, index, character);
                    AppendYear(value, culture, count, builder);
                    index += count - 1;
                    continue;
                }

                if (character == 'd' || character == 'D')
                {
                    var count = CountRepeated(formatCode, index, character);
                    AppendDay(value, culture, count, builder);
                    index += count - 1;
                    continue;
                }

                if (character == 'h' || character == 'H')
                {
                    var count = CountRepeated(formatCode, index, character);
                    AppendHour(value, culture, count, hasAmPm, builder);
                    index += count - 1;
                    continue;
                }

                if (character == 's' || character == 'S')
                {
                    var count = CountRepeated(formatCode, index, character);
                    AppendSecond(value, culture, count, builder);
                    index += count - 1;
                    continue;
                }

                if (character == 'm' || character == 'M')
                {
                    var count = CountRepeated(formatCode, index, character);
                    if (IsMinuteContext(formatCode, index, count))
                    {
                        AppendMinute(value, culture, count, builder);
                    }
                    else
                    {
                        AppendMonth(value, culture, count, builder);
                    }

                    index += count - 1;
                    continue;
                }

                builder.Append(character);
            }

            return builder.ToString();
        }

        internal static bool MatchesToken(string formatCode, int startIndex, string token)
        {
            if (startIndex + token.Length > formatCode.Length)
            {
                return false;
            }

            return string.Compare(formatCode, startIndex, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) == 0;
        }

        internal static int CountRepeated(string formatCode, int startIndex, char token)
        {
            var count = 1;
            for (var index = startIndex + 1; index < formatCode.Length; index++)
            {
                if (char.ToLowerInvariant(formatCode[index]) != char.ToLowerInvariant(token))
                {
                    break;
                }

                count++;
            }

            return count;
        }

        internal static bool IsMinuteContext(string formatCode, int startIndex, int count)
        {
            var previous = FindNeighborToken(formatCode, startIndex - 1, -1);
            var next = FindNeighborToken(formatCode, startIndex + count, 1);
            if (previous == ':' || next == ':')
            {
                return true;
            }

            if (previous == 'h' || previous == 'H' || previous == 's' || previous == 'S')
            {
                return true;
            }

            if (next == 'h' || next == 'H' || next == 's' || next == 'S')
            {
                return true;
            }

            return false;
        }

        internal static char FindNeighborToken(string formatCode, int startIndex, int direction)
        {
            var inQuote = false;
            if (direction < 0)
            {
                for (var index = startIndex; index >= 0; index--)
                {
                    var character = formatCode[index];
                    if (character == '"')
                    {
                        inQuote = !inQuote;
                        continue;
                    }

                    if (inQuote)
                    {
                        continue;
                    }

                    if (character == '\\' || character == '_' || character == '*')
                    {
                        index--;
                        continue;
                    }

                    if (char.IsWhiteSpace(character))
                    {
                        continue;
                    }

                    return character;
                }
            }
            else
            {
                for (var index = startIndex; index < formatCode.Length; index++)
                {
                    var character = formatCode[index];
                    if (character == '"')
                    {
                        inQuote = !inQuote;
                        continue;
                    }

                    if (inQuote)
                    {
                        continue;
                    }

                    if (character == '\\' || character == '_' || character == '*')
                    {
                        index++;
                        continue;
                    }

                    if (char.IsWhiteSpace(character))
                    {
                        continue;
                    }

                    return character;
                }
            }

            return '\0';
        }

        internal static bool IsElapsedToken(string token)
        {
            var normalized = token.Trim().ToLowerInvariant();
            return normalized == "h" || normalized == "hh" || normalized == "m" || normalized == "mm" || normalized == "s" || normalized == "ss";
        }

        internal static bool ContainsElapsedTimeToken(string formatCode)
        {
            for (var index = 0; index < formatCode.Length; index++)
            {
                if (formatCode[index] != '[')
                {
                    continue;
                }

                var endIndex = formatCode.IndexOf(']', index + 1);
                if (endIndex < 0)
                {
                    continue;
                }

                var token = formatCode.Substring(index + 1, endIndex - index - 1);
                if (IsElapsedToken(token))
                {
                    return true;
                }
            }

            return false;
        }

        internal static string FormatElapsedTimeValue(TimeSpan time, string formatCode, CultureInfo culture)
        {
            var builder = new StringBuilder(formatCode.Length + 8);
            var inQuote = false;

            for (var index = 0; index < formatCode.Length; index++)
            {
                var character = formatCode[index];
                if (character == '"')
                {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote)
                {
                    builder.Append(character);
                    continue;
                }

                if (character == '[')
                {
                    var endIndex = formatCode.IndexOf(']', index + 1);
                    if (endIndex > index)
                    {
                        var token = formatCode.Substring(index + 1, endIndex - index - 1).ToLowerInvariant();
                        if (token == "h")
                        {
                            builder.Append(((int)Math.Floor(time.TotalHours)).ToString(CultureInfo.InvariantCulture));
                        }
                        else if (token == "hh")
                        {
                            builder.Append(((int)Math.Floor(time.TotalHours)).ToString("00", CultureInfo.InvariantCulture));
                        }
                        else if (token == "m")
                        {
                            builder.Append(((int)Math.Floor(time.TotalMinutes)).ToString(CultureInfo.InvariantCulture));
                        }
                        else if (token == "mm")
                        {
                            builder.Append(((int)Math.Floor(time.TotalMinutes)).ToString("00", CultureInfo.InvariantCulture));
                        }
                        else if (token == "s")
                        {
                            builder.Append(((int)Math.Floor(time.TotalSeconds)).ToString(CultureInfo.InvariantCulture));
                        }
                        else if (token == "ss")
                        {
                            builder.Append(((int)Math.Floor(time.TotalSeconds)).ToString("00", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            builder.Append('[');
                            builder.Append(token);
                            builder.Append(']');
                        }

                        index = endIndex;
                        continue;
                    }
                }

                if (character == 'h' || character == 'H')
                {
                    var count = CountRepeated(formatCode, index, character);
                    if (count == 1)
                    {
                        builder.Append(time.Hours.ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        builder.Append(time.Hours.ToString("00", CultureInfo.InvariantCulture));
                    }

                    index += count - 1;
                    continue;
                }

                if (character == 'm' || character == 'M')
                {
                    var count = CountRepeated(formatCode, index, character);
                    if (count == 1)
                    {
                        builder.Append(time.Minutes.ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        builder.Append(time.Minutes.ToString("00", CultureInfo.InvariantCulture));
                    }

                    index += count - 1;
                    continue;
                }

                if (character == 's' || character == 'S')
                {
                    var count = CountRepeated(formatCode, index, character);
                    if (count == 1)
                    {
                        builder.Append(time.Seconds.ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        builder.Append(time.Seconds.ToString("00", CultureInfo.InvariantCulture));
                    }

                    index += count - 1;
                    continue;
                }

                if (character == '.' && index + 1 < formatCode.Length && formatCode[index + 1] == '0')
                {
                    var zeroCount = CountRepeated(formatCode, index + 1, '0');
                    builder.Append(culture.NumberFormat.NumberDecimalSeparator);
                    AppendFractionDigits(time.Milliseconds, zeroCount, builder);
                    index += zeroCount;
                    continue;
                }

                if (character == '\\')
                {
                    if (index + 1 < formatCode.Length)
                    {
                        index++;
                        builder.Append(formatCode[index]);
                    }

                    continue;
                }

                if (character == '_' || character == '*')
                {
                    index++;
                    continue;
                }

                builder.Append(character);
            }

            return builder.ToString();
        }

        private static void AppendYear(DateTime value, CultureInfo culture, int count, StringBuilder builder)
        {
            var year = value.Year;
            if (count <= 1)
            {
                builder.Append((year % 100).ToString(culture));
                return;
            }

            if (count == 2)
            {
                builder.Append((year % 100).ToString("00", culture));
                return;
            }

            builder.Append(year.ToString(new string('0', count), culture));
        }

        private static void AppendDay(DateTime value, CultureInfo culture, int count, StringBuilder builder)
        {
            if (count == 1)
            {
                builder.Append(value.Day.ToString(culture));
                return;
            }

            if (count == 2)
            {
                builder.Append(value.Day.ToString("00", culture));
                return;
            }

            if (count == 3)
            {
                builder.Append(culture.DateTimeFormat.GetAbbreviatedDayName(value.DayOfWeek));
                return;
            }

            builder.Append(culture.DateTimeFormat.GetDayName(value.DayOfWeek));
        }

        private static void AppendMonth(DateTime value, CultureInfo culture, int count, StringBuilder builder)
        {
            if (count == 1)
            {
                builder.Append(value.Month.ToString(culture));
                return;
            }

            if (count == 2)
            {
                builder.Append(value.Month.ToString("00", culture));
                return;
            }

            if (count == 3)
            {
                builder.Append(culture.DateTimeFormat.GetAbbreviatedMonthName(value.Month));
                return;
            }

            var monthName = culture.DateTimeFormat.GetMonthName(value.Month);
            if (count == 4)
            {
                builder.Append(monthName);
                return;
            }

            if (monthName.Length == 0)
            {
                builder.Append(value.Month.ToString(culture));
                return;
            }

            builder.Append(monthName[0]);
        }

        private static void AppendHour(DateTime value, CultureInfo culture, int count, bool hasAmPm, StringBuilder builder)
        {
            var hour = value.Hour;
            if (hasAmPm)
            {
                hour %= 12;
                if (hour == 0)
                {
                    hour = 12;
                }
            }

            if (count == 1)
            {
                builder.Append(hour.ToString(culture));
                return;
            }

            builder.Append(hour.ToString("00", culture));
        }

        private static void AppendMinute(DateTime value, CultureInfo culture, int count, StringBuilder builder)
        {
            if (count == 1)
            {
                builder.Append(value.Minute.ToString(culture));
                return;
            }

            builder.Append(value.Minute.ToString("00", culture));
        }

        private static void AppendSecond(DateTime value, CultureInfo culture, int count, StringBuilder builder)
        {
            if (count == 1)
            {
                builder.Append(value.Second.ToString(culture));
                return;
            }

            builder.Append(value.Second.ToString("00", culture));
        }

        private static bool TryAppendFractionalSeconds(DateTime value, string formatCode, CultureInfo culture, ref int index, StringBuilder builder)
        {
            if (index + 1 >= formatCode.Length || formatCode[index + 1] != '0')
            {
                return false;
            }

            var previous = FindNeighborToken(formatCode, index - 1, -1);
            if (previous != 's' && previous != 'S')
            {
                return false;
            }

            var zeroCount = CountRepeated(formatCode, index + 1, '0');
            builder.Append(culture.NumberFormat.NumberDecimalSeparator);
            AppendFractionDigits(value.Millisecond, zeroCount, builder);
            index += zeroCount;
            return true;
        }

        private static void AppendFractionDigits(int milliseconds, int zeroCount, StringBuilder builder)
        {
            var digits = milliseconds.ToString("000", CultureInfo.InvariantCulture);
            if (zeroCount <= 0)
            {
                return;
            }

            if (zeroCount == 1)
            {
                builder.Append(digits[0]);
                return;
            }

            if (zeroCount == 2)
            {
                builder.Append(digits[0]);
                builder.Append(digits[1]);
                return;
            }

            builder.Append(digits);
            for (var index = 3; index < zeroCount; index++)
            {
                builder.Append('0');
            }
        }

        private static string GetAmPmDesignator(DateTime value, CultureInfo culture, bool abbreviated)
        {
            var designator = value.Hour < 12 ? culture.DateTimeFormat.AMDesignator : culture.DateTimeFormat.PMDesignator;
            if (string.IsNullOrEmpty(designator))
            {
                designator = value.Hour < 12 ? "AM" : "PM";
            }

            if (!abbreviated)
            {
                return designator;
            }

            return designator.Substring(0, 1);
        }
    }
}
