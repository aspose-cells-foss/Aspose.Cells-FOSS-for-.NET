using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class ChartDateAxisLayout
    {
        public int MonthStep;
        public float RotationDeg;

        public static ChartDateAxisLayout Resolve(ParsedChart chart, DateSystem dateSystem, float availableWidth)
        {
            var result = new ChartDateAxisLayout();
            result.MonthStep = 1;
            result.RotationDeg = 0f;

            if (chart == null || !chart.HasDateCategoryAxis || chart.CategoryValues == null)
            {
                return result;
            }

            double minDate;
            double maxDate;
            if (!TryGetDateAxisRange(chart, out minDate, out maxDate))
            {
                return result;
            }

            var visibleCategoryCount = CountVisibleCategoryValues(chart);
            var monthCount = GetInclusiveMonthCount(minDate, maxDate, dateSystem);
            if (Math.Abs(chart.CategoryAxisTextRotationDeg) <= 0.1d)
            {
                return result;
            }

            var visibleTickCount = CountVisibleTickCount(monthCount, result.MonthStep);
            var slotWidth = availableWidth;
            if (visibleTickCount > 1)
            {
                slotWidth = availableWidth / (visibleTickCount - 1);
            }

            if (monthCount > visibleCategoryCount + 2 || slotWidth < 52f)
            {
                result.RotationDeg = (float)chart.CategoryAxisTextRotationDeg;
            }

            return result;
        }

        private static bool TryGetDateAxisRange(ParsedChart chart, out double minDate, out double maxDate)
        {
            minDate = double.MaxValue;
            maxDate = double.MinValue;
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

        private static int CountVisibleCategoryValues(ParsedChart chart)
        {
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

        private static int CountVisibleTickCount(int monthCount, int monthStep)
        {
            if (monthCount <= 0)
            {
                return 0;
            }

            if (monthStep <= 1)
            {
                return monthCount;
            }

            return (monthCount - 1) / monthStep + 1;
        }
    }
}
