using System;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class NumberFormatTextColor
    {
        internal static SKColor Resolve(object value, StyleValue style, SKColor fallback)
        {
            if (value == null || style == null)
            {
                return fallback;
            }

            var formatCode = NumberFormat.ResolveFormatCode(style.NumberFormat.Number, style.NumberFormat.Custom);
            if (string.IsNullOrWhiteSpace(formatCode))
            {
                return fallback;
            }

            var sections = DisplayTextFormatterSupport.ParseSections(formatCode);
            DisplayFormatSectionInfo selected = null;
            if (DisplayTextFormatterSupport.IsNumericValue(value))
            {
                bool useAbsoluteValue;
                selected = DisplayTextFormatterSupport.SelectNumericSection(sections,
                    DisplayTextFormatterSupport.ConvertToDouble(value), out useAbsoluteValue);
            }
            else if (value is DateTime)
            {
                selected = DisplayTextFormatterSupport.SelectDateTimeSection(sections);
            }
            else if (value is string)
            {
                selected = DisplayTextFormatterSupport.SelectTextSection(sections);
            }

            if (selected == null)
            {
                return fallback;
            }

            var section = selected.Raw;
            var inQuote = false;
            for (var index = 0; index < section.Length; index++)
            {
                var character = section[index];
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

                if (character != '[')
                {
                    continue;
                }

                var end = section.IndexOf(']', index + 1);
                if (end < 0)
                {
                    break;
                }

                var name = section.Substring(index + 1, end - index - 1);
                SKColor color;
                if (TryResolveNamedColor(name, out color))
                {
                    return color;
                }

                index = end;
            }

            return fallback;
        }

        private static bool TryResolveNamedColor(string name, out SKColor color)
        {
            switch (name.ToLowerInvariant())
            {
                case "black": color = SKColors.Black; return true;
                case "blue": color = SKColors.Blue; return true;
                case "cyan": color = SKColors.Cyan; return true;
                case "green": color = SKColors.Lime; return true;
                case "magenta": color = SKColors.Magenta; return true;
                case "red": color = SKColors.Red; return true;
                case "white": color = SKColors.White; return true;
                case "yellow": color = SKColors.Yellow; return true;
                default:
                    color = default(SKColor);
                    return false;
            }
        }
    }
}
