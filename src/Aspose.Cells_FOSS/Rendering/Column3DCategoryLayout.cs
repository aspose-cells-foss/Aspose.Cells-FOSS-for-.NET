using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class Column3DCategoryLayout
    {
        public readonly int SlotCount;
        public readonly List<double> Centers = new List<double>();
        public readonly List<string> Labels = new List<string>();

        public Column3DCategoryLayout(ParsedChart chart, int count,
            DateSystem dateSystem, CultureInfo culture)
        {
            SlotCount = count;
            double min;
            double max;
            if (chart.HasDateCategoryAxis
                && string.Equals(chart.CategoryBaseTimeUnit, "months", StringComparison.Ordinal)
                && ChartDateAxisMath.TryGetDateAxisRange(chart, out min, out max))
            {
                var first = DateSerialConverter.FromSerial(min, dateSystem);
                var last = DateSerialConverter.FromSerial(max, dateSystem);
                SlotCount = (last.Year - first.Year) * 12 + last.Month - first.Month + 1;
                for (var i = 0; i < count; i++)
                {
                    var slot = i;
                    if (i < chart.CategoryValues.Count && chart.CategoryValues[i].HasValue)
                    {
                        var date = DateSerialConverter.FromSerial(chart.CategoryValues[i].Value, dateSystem);
                        slot = (date.Year - first.Year) * 12 + date.Month - first.Month;
                    }
                    Centers.Add((slot + 0.5d) / SlotCount);
                }
                var start = ChartDateAxisMath.StartOfMonth(min, dateSystem);
                for (var i = 0; i < SlotCount; i++)
                {
                    Labels.Add(ChartXmlParser.FormatLabel(
                        ChartDateAxisMath.AddMonthsSerial(start, i, dateSystem),
                        chart.CategoryFormatCode, dateSystem, culture));
                }
                return;
            }
            for (var i = 0; i < count; i++)
            {
                Centers.Add((i + 0.5d) / count);
                Labels.Add(i < chart.Categories.Count ? chart.Categories[i] : string.Empty);
            }
        }
    }
}
