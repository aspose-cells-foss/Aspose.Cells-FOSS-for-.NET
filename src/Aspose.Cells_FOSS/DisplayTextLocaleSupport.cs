using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using System.Text;

namespace Aspose.Cells_FOSS
{
    internal static class DisplayTextLocaleSupport
    {
        internal static string ApplyLocaleDirectives(string section, CultureInfo fallbackCulture, out CultureInfo sectionCulture)
        {
            sectionCulture = fallbackCulture;
            var builder = new StringBuilder(section.Length + 16);
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
                    if (endIndex > index)
                    {
                        var token = section.Substring(index + 1, endIndex - index - 1);
                        string replacement;
                        CultureInfo resolvedCulture;
                        if (TryResolveLocaleDirective(token, sectionCulture, out replacement, out resolvedCulture))
                        {
                            builder.Append(replacement);
                            sectionCulture = resolvedCulture;
                            index = endIndex;
                            continue;
                        }
                    }
                }

                builder.Append(character);
            }

            return builder.ToString();
        }

        private static bool TryResolveLocaleDirective(string token, CultureInfo fallbackCulture, out string replacement, out CultureInfo resolvedCulture)
        {
            replacement = string.Empty;
            resolvedCulture = fallbackCulture;
            if (string.IsNullOrWhiteSpace(token) || token[0] != '$')
            {
                return false;
            }

            var dashIndex = token.LastIndexOf('-');
            if (dashIndex <= 0 || dashIndex >= token.Length - 1)
            {
                return false;
            }

            var symbol = token.Substring(1, dashIndex - 1);
            var localeCode = token.Substring(dashIndex + 1);
            resolvedCulture = ResolveCulture(localeCode, fallbackCulture);

            if (string.Equals(localeCode, "F800", StringComparison.OrdinalIgnoreCase))
            {
                replacement = fallbackCulture.DateTimeFormat.LongDatePattern;
                resolvedCulture = fallbackCulture;
                return true;
            }

            if (string.Equals(localeCode, "F400", StringComparison.OrdinalIgnoreCase))
            {
                replacement = fallbackCulture.DateTimeFormat.LongTimePattern;
                resolvedCulture = fallbackCulture;
                return true;
            }

            if (!string.IsNullOrEmpty(symbol))
            {
                replacement = QuoteLiteral(symbol);
            }

            return true;
        }

        private static CultureInfo ResolveCulture(string localeCode, CultureInfo fallbackCulture)
        {
            if (string.IsNullOrWhiteSpace(localeCode))
            {
                return fallbackCulture;
            }

            int lcid;
            if (!int.TryParse(localeCode, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out lcid))
            {
                return fallbackCulture;
            }

            try
            {
                return CultureInfo.GetCultureInfo(lcid);
            }
            catch (CultureNotFoundException)
            {
                return fallbackCulture;
            }
        }

        private static string QuoteLiteral(string value)
        {
            var builder = new StringBuilder(value.Length + 2);
            builder.Append('"');
            for (var index = 0; index < value.Length; index++)
            {
                if (value[index] == '"')
                {
                    builder.Append('"');
                }

                builder.Append(value[index]);
            }

            builder.Append('"');
            return builder.ToString();
        }
    }
}
