using System;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartDateAxisMath
    {
        public static bool TryGetDateAxisPlotRange(ParsedChart chart, DateSystem dateSystem, out double plotMinDate, out double plotMaxDate)
        {
            plotMinDate = 0d;
            plotMaxDate = 0d;

            double minDate;
            double maxDate;
            if (!TryGetDateAxisRange(chart, out minDate, out maxDate))
            {
                return false;
            }

            if (chart == null || !chart.CrossBetween)
            {
                plotMinDate = minDate;
                plotMaxDate = maxDate;
                return true;
            }

            if (string.Equals(chart.CategoryBaseTimeUnit, "months", StringComparison.Ordinal))
            {
                var firstTick = StartOfMonth(minDate, dateSystem);
                var lastTick = StartOfMonth(maxDate, dateSystem);
                var previousTick = AddMonthsSerial(firstTick, -1, dateSystem);
                var nextTick = AddMonthsSerial(lastTick, 1, dateSystem);
                plotMinDate = Midpoint(previousTick, firstTick);
                plotMaxDate = Midpoint(lastTick, nextTick);
                return plotMaxDate > plotMinDate;
            }

            if (chart.CategoryValues != null && chart.CategoryValues.Count > 1)
            {
                double firstValue;
                double secondValue;
                double previousValue;
                double lastValue;
                if (TryGetFirstTwoCategoryValues(chart, out firstValue, out secondValue)
                    && TryGetLastTwoCategoryValues(chart, out previousValue, out lastValue))
                {
                    plotMinDate = firstValue - (secondValue - firstValue) * 0.5d;
                    plotMaxDate = lastValue + (lastValue - previousValue) * 0.5d;
                    return plotMaxDate > plotMinDate;
                }
            }

            plotMinDate = minDate;
            plotMaxDate = maxDate;
            return true;
        }

        public static bool TryGetDateAxisRange(ParsedChart chart, out double minDate, out double maxDate)
        {
            minDate = double.MaxValue;
            maxDate = double.MinValue;
            if (chart == null || chart.CategoryValues == null)
            {
                return false;
            }

            for (var i = 0; i < chart.CategoryValues.Count; i++)
            {
                if (!chart.CategoryValues[i].HasValue)
                {
                    continue;
                }

                var value = chart.CategoryValues[i].Value;
                if (value < minDate)
                {
                    minDate = value;
                }

                if (value > maxDate)
                {
                    maxDate = value;
                }
            }

            if (minDate == double.MaxValue || maxDate == double.MinValue)
            {
                minDate = 0d;
                maxDate = 0d;
                return false;
            }

            return true;
        }

        private static bool TryGetFirstTwoCategoryValues(ParsedChart chart, out double firstValue, out double secondValue)
        {
            firstValue = 0d;
            secondValue = 0d;
            var found = 0;
            for (var i = 0; i < chart.CategoryValues.Count; i++)
            {
                if (!chart.CategoryValues[i].HasValue)
                {
                    continue;
                }

                if (found == 0)
                {
                    firstValue = chart.CategoryValues[i].Value;
                    found = 1;
                }
                else
                {
                    secondValue = chart.CategoryValues[i].Value;
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetLastTwoCategoryValues(ParsedChart chart, out double previousValue, out double lastValue)
        {
            previousValue = 0d;
            lastValue = 0d;
            var foundLast = false;
            for (var i = chart.CategoryValues.Count - 1; i >= 0; i--)
            {
                if (!chart.CategoryValues[i].HasValue)
                {
                    continue;
                }

                if (!foundLast)
                {
                    lastValue = chart.CategoryValues[i].Value;
                    foundLast = true;
                }
                else
                {
                    previousValue = chart.CategoryValues[i].Value;
                    return true;
                }
            }

            return false;
        }

        private static double Midpoint(double a, double b)
        {
            return a + (b - a) * 0.5d;
        }

        public static double StartOfMonth(double serial, DateSystem dateSystem)
        {
            var date = DateSerialConverter.FromSerial(serial, dateSystem);
            var first = new DateTime(date.Year, date.Month, 1);
            return DateSerialConverter.ToSerial(first, dateSystem);
        }

        public static double AddMonthsSerial(double serial, int months, DateSystem dateSystem)
        {
            var date = DateSerialConverter.FromSerial(serial, dateSystem);
            var next = new DateTime(date.Year, date.Month, 1).AddMonths(months);
            return DateSerialConverter.ToSerial(next, dateSystem);
        }

        public static float DateCategoryX(double serial, double minDate, double maxDate, SKRect plot)
        {
            if (maxDate <= minDate)
            {
                return plot.MidX;
            }

            var t = (serial - minDate) / (maxDate - minDate);
            if (t < 0d)
            {
                t = 0d;
            }
            else if (t > 1d)
            {
                t = 1d;
            }

            return plot.Left + (float)t * plot.Width;
        }

        public static double NormalizePointSerial(ParsedChart chart, double serial, DateSystem dateSystem)
        {
            if (chart == null || !chart.HasDateCategoryAxis)
            {
                return serial;
            }

            if (!string.Equals(chart.CategoryBaseTimeUnit, "months", StringComparison.Ordinal))
            {
                return serial;
            }

            double minDate;
            double maxDate;
            if (!TryGetDateAxisRange(chart, out minDate, out maxDate))
            {
                return serial;
            }

            var visibleCategoryCount = CountVisibleCategoryValues(chart);
            var monthCount = GetInclusiveMonthCount(minDate, maxDate, dateSystem);
            if (visibleCategoryCount > 0 && monthCount <= visibleCategoryCount + 2)
            {
                return StartOfMonth(serial, dateSystem);
            }

            return serial;
        }

        private static int CountVisibleCategoryValues(ParsedChart chart)
        {
            if (chart == null || chart.CategoryValues == null)
            {
                return 0;
            }

            var count = 0;
            for (var i = 0; i < chart.CategoryValues.Count; i++)
            {
                if (chart.CategoryValues[i].HasValue)
                {
                    count++;
                }
            }

            return count;
        }

        private static int GetInclusiveMonthCount(double minDate, double maxDate, DateSystem dateSystem)
        {
            var min = DateSerialConverter.FromSerial(minDate, dateSystem);
            var max = DateSerialConverter.FromSerial(maxDate, dateSystem);
            return (max.Year - min.Year) * 12 + max.Month - min.Month + 1;
        }
    }
}
