using System.IO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Aspose.Cells_FOSS
{
    internal static class DisplayTextFormatterSupport
    {
        internal static DisplayFormatSectionInfo SelectNumericSection(IReadOnlyList<DisplayFormatSectionInfo> sections, double numericValue, out bool useAbsoluteValue)
        {
            useAbsoluteValue = false;
            var numericSectionCount = sections.Count < 3 ? sections.Count : 3;
            if (numericSectionCount == 0)
            {
                return null;
            }

            var hasCondition = false;
            DisplayFormatSectionInfo fallbackSection = null;
            for (var index = 0; index < numericSectionCount; index++)
            {
                var section = sections[index];
                if (section.HasCondition)
                {
                    hasCondition = true;
                    if (EvaluateCondition(section.ConditionOperator, numericValue, section.ConditionValue))
                    {
                        useAbsoluteValue = ShouldUseAbsoluteValue(section.Raw, numericValue);
                        return section;
                    }
                }
                else if (fallbackSection == null)
                {
                    fallbackSection = section;
                }
            }

            if (hasCondition && fallbackSection != null)
            {
                useAbsoluteValue = ShouldUseAbsoluteValue(fallbackSection.Raw, numericValue);
                return fallbackSection;
            }

            if (numericSectionCount == 1)
            {
                useAbsoluteValue = ShouldUseAbsoluteValue(sections[0].Raw, numericValue);
                return sections[0];
            }

            if (numericSectionCount == 2)
            {
                if (numericValue < 0)
                {
                    useAbsoluteValue = ShouldUseAbsoluteValue(sections[1].Raw, numericValue);
                    return sections[1];
                }

                useAbsoluteValue = ShouldUseAbsoluteValue(sections[0].Raw, numericValue);
                return sections[0];
            }

            if (numericValue > 0)
            {
                useAbsoluteValue = ShouldUseAbsoluteValue(sections[0].Raw, numericValue);
                return sections[0];
            }

            if (numericValue < 0)
            {
                useAbsoluteValue = ShouldUseAbsoluteValue(sections[1].Raw, numericValue);
                return sections[1];
            }

            useAbsoluteValue = false;
            return sections[2];
        }

        internal static DisplayFormatSectionInfo SelectTextSection(IReadOnlyList<DisplayFormatSectionInfo> sections)
        {
            if (sections.Count >= 4)
            {
                return sections[3];
            }

            for (var index = 0; index < sections.Count; index++)
            {
                var section = StripDirectiveBrackets(sections[index].Raw, false);
                if (section.IndexOf('@') >= 0)
                {
                    return sections[index];
                }
            }

            return null;
        }

        internal static DisplayFormatSectionInfo SelectDateTimeSection(IReadOnlyList<DisplayFormatSectionInfo> sections)
        {
            var dateSectionCount = sections.Count < 3 ? sections.Count : 3;
            for (var index = 0; index < dateSectionCount; index++)
            {
                if (!string.IsNullOrWhiteSpace(sections[index].Raw))
                {
                    return sections[index];
                }
            }

            return null;
        }

        internal static List<DisplayFormatSectionInfo> ParseSections(string formatCode)
        {
            var rawSections = SplitSections(formatCode);
            var sections = new List<DisplayFormatSectionInfo>(rawSections.Count);
            for (var index = 0; index < rawSections.Count; index++)
            {
                var rawSection = rawSections[index];
                var section = new DisplayFormatSectionInfo
                {
                    Raw = rawSection,
                };

                string conditionOperator;
                double conditionValue;
                if (TryParseSectionCondition(rawSection, out conditionOperator, out conditionValue))
                {
                    section.HasCondition = true;
                    section.ConditionOperator = conditionOperator;
                    section.ConditionValue = conditionValue;
                }

                sections.Add(section);
            }

            return sections;
        }

        internal static List<string> SplitSections(string formatCode)
        {
            var sections = new List<string>();
            var builder = new StringBuilder();
            var inQuote = false;

            for (var index = 0; index < formatCode.Length; index++)
            {
                var character = formatCode[index];
                if (character == '"')
                {
                    builder.Append(character);
                    inQuote = !inQuote;
                    continue;
                }

                if (!inQuote)
                {
                    if (character == '\\' && index + 1 < formatCode.Length)
                    {
                        builder.Append(character);
                        index++;
                        builder.Append(formatCode[index]);
                        continue;
                    }

                    if (character == ';')
                    {
                        sections.Add(builder.ToString());
                        builder.Clear();
                        continue;
                    }
                }

                builder.Append(character);
            }

            sections.Add(builder.ToString());
            return sections;
        }

        internal static bool TryParseSectionCondition(string section, out string conditionOperator, out double conditionValue)
        {
            conditionOperator = string.Empty;
            conditionValue = 0;
            var inQuote = false;

            for (var index = 0; index < section.Length; index++)
            {
                var character = section[index];
                if (character == '"')
                {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote || character != '[')
                {
                    continue;
                }

                var endIndex = section.IndexOf(']', index + 1);
                if (endIndex < 0)
                {
                    break;
                }

                var token = section.Substring(index + 1, endIndex - index - 1).Trim();
                if (TryParseConditionToken(token, out conditionOperator, out conditionValue))
                {
                    return true;
                }

                index = endIndex;
            }

            return false;
        }

        internal static bool TryParseConditionToken(string token, out string conditionOperator, out double conditionValue)
        {
            conditionOperator = string.Empty;
            conditionValue = 0;
            if (string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            var operators = new string[] { ">=", "<=", "<>", ">", "<", "=" };
            for (var index = 0; index < operators.Length; index++)
            {
                var candidate = operators[index];
                if (token.StartsWith(candidate, StringComparison.Ordinal))
                {
                    var numberPart = token.Substring(candidate.Length).Trim();
                    double parsedValue;
                    if (double.TryParse(numberPart, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out parsedValue))
                    {
                        conditionOperator = candidate;
                        conditionValue = parsedValue;
                        return true;
                    }
                }
            }

            return false;
        }
        internal static bool EvaluateCondition(string conditionOperator, double numericValue, double conditionValue)
        {
            switch (conditionOperator)
            {
                case ">":
                    return numericValue > conditionValue;
                case ">=":
                    return numericValue >= conditionValue;
                case "<":
                    return numericValue < conditionValue;
                case "<=":
                    return numericValue <= conditionValue;
                case "=":
                    return Math.Abs(numericValue - conditionValue) < 1E-12;
                case "<>":
                    return Math.Abs(numericValue - conditionValue) >= 1E-12;
                default:
                    return false;
            }
        }

        internal static bool ShouldUseAbsoluteValue(string section, double numericValue)
        {
            if (numericValue >= 0)
            {
                return false;
            }

            var sanitizedSection = SanitizeNumericSection(section);
            return sanitizedSection.IndexOf('-') < 0;
        }

        internal static string SanitizeNumericSection(string section)
        {
            var withoutDirectives = StripDirectiveBrackets(section, false);
            var builder = new StringBuilder(withoutDirectives.Length);
            var inQuote = false;

            for (var index = 0; index < withoutDirectives.Length; index++)
            {
                var character = withoutDirectives[index];
                if (character == '"')
                {
                    builder.Append(character);
                    inQuote = !inQuote;
                    continue;
                }

                if (!inQuote)
                {
                    if (character == '_')
                    {
                        index++;
                        continue;
                    }

                    if (character == '*')
                    {
                        index++;
                        continue;
                    }

                    if (character == '\\')
                    {
                        if (index + 1 < withoutDirectives.Length)
                        {
                            index++;
                            builder.Append(withoutDirectives[index]);
                        }

                        continue;
                    }

                    if (character == '?')
                    {
                        builder.Append('#');
                        continue;
                    }

                    if (character == '[' || character == ']')
                    {
                        continue;
                    }
                }

                builder.Append(character);
            }

            return builder.ToString().Trim();
        }

        internal static string StripDirectiveBrackets(string section, bool preserveElapsedTokens)
        {
            var builder = new StringBuilder(section.Length);
            var inQuote = false;

            for (var index = 0; index < section.Length; index++)
            {
                var character = section[index];
                if (character == '"')
                {
                    builder.Append(character);
                    inQuote = !inQuote;
                    continue;
                }

                if (!inQuote && character == '[')
                {
                    var endIndex = section.IndexOf(']', index + 1);
                    if (endIndex < 0)
                    {
                        continue;
                    }

                    var token = section.Substring(index + 1, endIndex - index - 1);
                    if (preserveElapsedTokens && DisplayTextDateFormatSupport.IsElapsedToken(token))
                    {
                        builder.Append('[');
                        builder.Append(token);
                        builder.Append(']');
                    }

                    index = endIndex;
                    continue;
                }

                builder.Append(character);
            }

            return builder.ToString();
        }

        internal static bool ContainsNumericPlaceholder(string pattern)
        {
            for (var index = 0; index < pattern.Length; index++)
            {
                var character = pattern[index];
                if (character == '0' || character == '#' || character == '?' || character == '.' || character == '%' || character == 'E' || character == 'e' || character == '/')
                {
                    return true;
                }
            }

            return false;
        }

        internal static string ExpandSectionPattern(string pattern, string valueText, bool replaceTextPlaceholder)
        {
            var builder = new StringBuilder(pattern.Length + valueText.Length);
            var inQuote = false;

            for (var index = 0; index < pattern.Length; index++)
            {
                var character = pattern[index];
                if (character == '"')
                {
                    inQuote = !inQuote;
                    continue;
                }

                if (!inQuote)
                {
                    if (character == '@' && replaceTextPlaceholder)
                    {
                        builder.Append(valueText);
                        continue;
                    }

                    if (character == '_')
                    {
                        index++;
                        continue;
                    }

                    if (character == '*')
                    {
                        index++;
                        continue;
                    }

                    if (character == '\\')
                    {
                        if (index + 1 < pattern.Length)
                        {
                            index++;
                            builder.Append(pattern[index]);
                        }

                        continue;
                    }
                }

                builder.Append(character);
            }

            return builder.ToString();
        }

        internal static bool IsNumericValue(object value)
        {
            if (value is byte || value is short || value is int || value is long)
            {
                return true;
            }

            if (value is float || value is double || value is decimal)
            {
                return true;
            }

            return false;
        }

        internal static double ConvertToDouble(object value)
        {
            if (value is byte) return (byte)value;
            if (value is short) return (short)value;
            if (value is int) return (int)value;
            if (value is long) return (long)value;
            if (value is float) return (float)value;
            if (value is double) return (double)value;
            if (value is decimal) return (double)(decimal)value;
            throw new InvalidOperationException("Value is not numeric.");
        }

        internal static object GetAbsoluteNumericValue(object value)
        {
            if (value is byte) return (byte)value;
            if (value is short) return Math.Abs((short)value);
            if (value is int) return Math.Abs((int)value);
            if (value is long) return Math.Abs((long)value);
            if (value is float) return Math.Abs((float)value);
            if (value is double) return Math.Abs((double)value);
            if (value is decimal) return Math.Abs((decimal)value);
            return value;
        }
    }
}
