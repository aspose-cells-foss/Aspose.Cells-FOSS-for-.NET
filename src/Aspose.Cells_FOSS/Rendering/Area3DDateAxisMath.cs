using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class Area3DDateAxisMath
    {
        public static double CategoryPosition(ParsedChart chart, int index, int count, DateSystem dateSystem)
        {
            if (chart != null
                && chart.HasDateCategoryAxis
                && index >= 0
                && index < chart.CategoryValues.Count
                && chart.CategoryValues[index].HasValue)
            {
                double minDate;
                double maxDate;
                if (ChartDateAxisMath.TryGetDateAxisPlotRange(chart, dateSystem, out minDate, out maxDate)
                    && maxDate > minDate)
                {
                    var serial = ChartDateAxisMath.NormalizePointSerial(chart, chart.CategoryValues[index].Value, dateSystem);
                    var position = (serial - minDate) / (maxDate - minDate);
                    position = Clamp(position);
                    return chart.CategoryAxisMinMax ? position : 1d - position;
                }
            }

            if (count <= 1)
            {
                return 0.5d;
            }

            if (chart != null && chart.CrossBetween)
            {
                return (index + 0.5d) / count;
            }

            return (double)index / (count - 1);
        }

        public static List<Tuple<double, string>> BuildMonthlyTicks(ParsedChart chart, DateSystem dateSystem, CultureInfo culture)
        {
            var result = new List<Tuple<double, string>>();
            if (chart == null
                || !chart.HasDateCategoryAxis
                || !string.Equals(chart.CategoryBaseTimeUnit, "months", StringComparison.Ordinal))
            {
                return result;
            }

            double minDate;
            double maxDate;
            double plotMinDate;
            double plotMaxDate;
            if (!ChartDateAxisMath.TryGetDateAxisRange(chart, out minDate, out maxDate)
                || !ChartDateAxisMath.TryGetDateAxisPlotRange(chart, dateSystem, out plotMinDate, out plotMaxDate))
            {
                return result;
            }

            var tick = ChartDateAxisMath.StartOfMonth(minDate, dateSystem);
            var last = ChartDateAxisMath.StartOfMonth(maxDate, dateSystem);
            var formatCode = string.IsNullOrEmpty(chart.CategoryFormatCode) ? "mmm-yy" : chart.CategoryFormatCode;
            while (tick <= last)
            {
                var position = Clamp((tick - plotMinDate) / (plotMaxDate - plotMinDate));
                if (!chart.CategoryAxisMinMax)
                {
                    position = 1d - position;
                }

                var label = ChartXmlParser.FormatLabel(tick, formatCode, dateSystem, culture);
                result.Add(Tuple.Create(position, label != null ? label.TrimEnd() : string.Empty));
                tick = ChartDateAxisMath.AddMonthsSerial(tick, 1, dateSystem);
            }

            return result;
        }

        public static void SortByPosition(List<double> positions, List<SKPoint> frontPoints, List<SKPoint> backPoints)
        {
            for (var i = 1; i < positions.Count; i++)
            {
                var position = positions[i];
                var front = frontPoints[i];
                var back = backPoints[i];
                var j = i - 1;
                while (j >= 0 && positions[j] > position)
                {
                    positions[j + 1] = positions[j];
                    frontPoints[j + 1] = frontPoints[j];
                    backPoints[j + 1] = backPoints[j];
                    j--;
                }

                positions[j + 1] = position;
                frontPoints[j + 1] = front;
                backPoints[j + 1] = back;
            }
        }

        private static double Clamp(double value)
        {
            if (value < 0d)
            {
                return 0d;
            }

            if (value > 1d)
            {
                return 1d;
            }

            return value;
        }
    }
}
