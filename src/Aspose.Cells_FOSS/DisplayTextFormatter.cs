using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    internal static class DisplayTextFormatter
    {
        internal static string FormatStringValue(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            if (value is string)
            {
                return (string)value;
            }

            if (value is bool)
            {
                return (bool)value ? "TRUE" : "FALSE";
            }

            if (value is DateTime)
            {
                return FormatRawDateTimeValue((DateTime)value);
            }

            if (value is byte)
            {
                return ((byte)value).ToString(CultureInfo.InvariantCulture);
            }

            if (value is short)
            {
                return ((short)value).ToString(CultureInfo.InvariantCulture);
            }

            if (value is int)
            {
                return ((int)value).ToString(CultureInfo.InvariantCulture);
            }

            if (value is long)
            {
                return ((long)value).ToString(CultureInfo.InvariantCulture);
            }

            if (value is float)
            {
                return ((float)value).ToString(null, CultureInfo.InvariantCulture);
            }

            if (value is double)
            {
                return ((double)value).ToString(null, CultureInfo.InvariantCulture);
            }

            if (value is decimal)
            {
                return ((decimal)value).ToString(null, CultureInfo.InvariantCulture);
            }

            var formattable = value as IFormattable;
            if (formattable != null)
            {
                return formattable.ToString(null, CultureInfo.InvariantCulture);
            }

            return value.ToString() ?? string.Empty;
        }

        internal static string FormatDisplayValue(object value, StyleValue style, CultureInfo workbookCulture)
        {
            if (value == null)
            {
                return string.Empty;
            }

            if (value is string)
            {
                return FormatTextValue((string)value, style);
            }

            if (value is bool)
            {
                return (bool)value ? "TRUE" : "FALSE";
            }

            if (value is DateTime)
            {
                return FormatDateTimeValue((DateTime)value, style, workbookCulture);
            }

            if (DisplayTextFormatterSupport.IsNumericValue(value))
            {
                return FormatNumericValue(value, style, workbookCulture);
            }

            var formattableValue = value as IFormattable;
            if (formattableValue != null)
            {
                return formattableValue.ToString(null, workbookCulture);
            }

            return value.ToString() ?? string.Empty;
        }

        private static string FormatTextValue(string value, StyleValue style)
        {
            var formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);
            if (string.IsNullOrWhiteSpace(formatCode) || string.Equals(formatCode, "General", StringComparison.Ordinal))
            {
                return value;
            }

            // Excel text formats can have a dedicated fourth section; only that section
            // should shape the displayed text without mutating the underlying value.
            var sections = DisplayTextFormatterSupport.ParseSections(formatCode);
            var textSection = DisplayTextFormatterSupport.SelectTextSection(sections);
            if (textSection == null)
            {
                return value;
            }

            var pattern = DisplayTextFormatterSupport.StripDirectiveBrackets(textSection.Raw, false);
            var formatted = DisplayTextFormatterSupport.ExpandSectionPattern(pattern, value, true);
            if (formatted.Length == 0)
            {
                return value;
            }

            return formatted;
        }

        private static string FormatNumericValue(object value, StyleValue style, CultureInfo workbookCulture)
        {
            var formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);
            if (string.IsNullOrWhiteSpace(formatCode) || string.Equals(formatCode, "General", StringComparison.Ordinal))
            {
                return FormatStringValue(value);
            }

            // Numeric formats can split positive, negative, zero, and text behavior into
            // separate sections, so the display path first selects the matching section.
            var sections = DisplayTextFormatterSupport.ParseSections(formatCode);
            if (sections.Count == 0)
            {
                return FormatStringValue(value);
            }

            var numericValue = DisplayTextFormatterSupport.ConvertToDouble(value);
            bool useAbsoluteValue;
            var selectedSection = DisplayTextFormatterSupport.SelectNumericSection(sections, numericValue, out useAbsoluteValue);
            if (selectedSection == null || string.IsNullOrWhiteSpace(selectedSection.Raw))
            {
                return FormatStringValue(value);
            }

            string fractionResult;
            // Fraction placeholders are not directly compatible with .NET numeric format
            // strings, so they are handled separately before standard formatting.
            if (TryFormatFraction(numericValue, selectedSection.Raw, useAbsoluteValue, out fractionResult))
            {
                return fractionResult;
            }

            CultureInfo sectionCulture;
            var localizedSection = DisplayTextLocaleSupport.ApplyLocaleDirectives(selectedSection.Raw, workbookCulture, out sectionCulture);
            var sanitizedSection = DisplayTextFormatterSupport.SanitizeNumericSection(localizedSection);
            if (string.IsNullOrWhiteSpace(sanitizedSection))
            {
                return FormatStringValue(value);
            }

            if (!DisplayTextFormatterSupport.ContainsNumericPlaceholder(sanitizedSection))
            {
                var literal = DisplayTextFormatterSupport.ExpandSectionPattern(sanitizedSection, string.Empty, false);
                if (literal.Length > 0)
                {
                    return literal;
                }

                return FormatStringValue(value);
            }

            try
            {
                var formattedValue = value;
                if (useAbsoluteValue && numericValue < 0)
                {
                    formattedValue = DisplayTextFormatterSupport.GetAbsoluteNumericValue(value);
                }

                var formattable = formattedValue as IFormattable;
                if (formattable != null)
                {
                    return formattable.ToString(sanitizedSection, sectionCulture);
                }
            }
            catch (FormatException)
            {
                return FormatStringValue(value);
            }

            return FormatStringValue(value);
        }

        private static string FormatDateTimeValue(DateTime value, StyleValue style, CultureInfo workbookCulture)
        {
            var formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);
            if (string.IsNullOrWhiteSpace(formatCode) || string.Equals(formatCode, "General", StringComparison.Ordinal))
            {
                return FormatRawDateTimeValue(value);
            }

            // Date/time display uses the first section that still carries date tokens
            // after locale and directive cleanup.
            var sections = DisplayTextFormatterSupport.ParseSections(formatCode);
            var section = DisplayTextFormatterSupport.SelectDateTimeSection(sections);
            if (section == null || string.IsNullOrWhiteSpace(section.Raw))
            {
                return FormatRawDateTimeValue(value);
            }

            CultureInfo sectionCulture;
            var localizedSection = DisplayTextLocaleSupport.ApplyLocaleDirectives(section.Raw, workbookCulture, out sectionCulture);
            var sectionFormat = DisplayTextFormatterSupport.StripDirectiveBrackets(localizedSection, true);
            // Elapsed-time formats are duration-style output, not calendar rendering.
            if (DisplayTextDateFormatSupport.ContainsElapsedTimeToken(sectionFormat))
            {
                return DisplayTextDateFormatSupport.FormatElapsedTimeValue(value.TimeOfDay, sectionFormat, sectionCulture);
            }

            if (string.IsNullOrWhiteSpace(sectionFormat))
            {
                return FormatRawDateTimeValue(value);
            }

            try
            {
                return DisplayTextDateFormatSupport.FormatDateTimeValue(value, sectionFormat, sectionCulture);
            }
            catch (FormatException)
            {
                return FormatRawDateTimeValue(value);
            }
        }

        private static string FormatRawDateTimeValue(DateTime value)
        {
            if (value.TimeOfDay == TimeSpan.Zero)
            {
                return value.ToString("M/d/yyyy", CultureInfo.InvariantCulture);
            }

            return value.ToString("M/d/yyyy H:mm", CultureInfo.InvariantCulture);
        }

        private static bool TryFormatFraction(double numericValue, string section, bool useAbsoluteValue, out string result)
        {
            result = string.Empty;
            var sanitizedSection = DisplayTextFormatterSupport.SanitizeNumericSection(section);
            if (sanitizedSection.IndexOf('/') < 0)
            {
                return false;
            }

            var slashIndex = sanitizedSection.IndexOf('/');
            if (slashIndex <= 0)
            {
                return false;
            }

            var denominatorDigits = 0;
            for (var index = slashIndex + 1; index < sanitizedSection.Length; index++)
            {
                var character = sanitizedSection[index];
                if (character == '#' || character == '0')
                {
                    denominatorDigits++;
                    continue;
                }

                if (char.IsWhiteSpace(character))
                {
                    continue;
                }

                break;
            }

            if (denominatorDigits <= 0)
            {
                return false;
            }

            var absoluteValue = Math.Abs(numericValue);
            var wholePart = (long)Math.Floor(absoluteValue);
            var fractionalPart = absoluteValue - wholePart;
            if (fractionalPart < 1E-12)
            {
                result = FormatWholeFractionResult(wholePart, useAbsoluteValue, numericValue);
                return true;
            }

            var maxDenominator = 1;
            for (var index = 0; index < denominatorDigits; index++)
            {
                maxDenominator *= 10;
            }

            // Match Excel's visible fraction behavior by finding the closest rational
            // value that fits within the placeholder width from the format section.
            maxDenominator -= 1;
            var bestNumerator = 0;
            var bestDenominator = 1;
            var bestError = double.MaxValue;

            for (var denominator = 1; denominator <= maxDenominator; denominator++)
            {
                var numerator = (int)Math.Round(fractionalPart * denominator, MidpointRounding.AwayFromZero);
                if (numerator == 0)
                {
                    continue;
                }

                if (numerator > denominator)
                {
                    numerator = denominator;
                }

                var candidate = (double)numerator / denominator;
                var error = Math.Abs(fractionalPart - candidate);
                if (error < bestError)
                {
                    bestError = error;
                    bestNumerator = numerator;
                    bestDenominator = denominator;
                }
            }

            if (bestNumerator == 0)
            {
                result = FormatWholeFractionResult(wholePart, useAbsoluteValue, numericValue);
                return true;
            }

            var greatestCommonDivisor = GreatestCommonDivisor(bestNumerator, bestDenominator);
            bestNumerator /= greatestCommonDivisor;
            bestDenominator /= greatestCommonDivisor;

            if (bestNumerator == bestDenominator)
            {
                wholePart++;
                bestNumerator = 0;
            }

            var prefix = string.Empty;
            if (!useAbsoluteValue && numericValue < 0)
            {
                prefix = "-";
            }

            if (bestNumerator == 0)
            {
                result = prefix + wholePart.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            if (wholePart == 0)
            {
                result = prefix + bestNumerator.ToString(CultureInfo.InvariantCulture) + "/" + bestDenominator.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            result = prefix + wholePart.ToString(CultureInfo.InvariantCulture) + " " + bestNumerator.ToString(CultureInfo.InvariantCulture) + "/" + bestDenominator.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static string FormatWholeFractionResult(long wholePart, bool useAbsoluteValue, double numericValue)
        {
            if (!useAbsoluteValue && numericValue < 0)
            {
                return "-" + wholePart.ToString(CultureInfo.InvariantCulture);
            }

            return wholePart.ToString(CultureInfo.InvariantCulture);
        }

        private static int GreatestCommonDivisor(int left, int right)
        {
            var first = Math.Abs(left);
            var second = Math.Abs(right);
            while (second != 0)
            {
                var remainder = first % second;
                first = second;
                second = remainder;
            }

            if (first == 0)
            {
                return 1;
            }

            return first;
        }
    }
}
