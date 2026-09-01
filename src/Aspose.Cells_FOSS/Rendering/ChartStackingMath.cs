using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartStackingMath
    {
        public static void ComputeDataRange(ParsedChart chart, out double minValue, out double maxValue)
        {
            minValue = double.MaxValue;
            maxValue = double.MinValue;

            if (chart == null)
            {
                return;
            }

            if (!chart.IsStacked && !chart.IsPercentStacked)
            {
                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    for (var i = 0; i < series.Values.Count; i++)
                    {
                        if (!series.Values[i].HasValue)
                        {
                            continue;
                        }

                        var value = series.Values[i].Value;
                        if (value < minValue)
                        {
                            minValue = value;
                        }

                        if (value > maxValue)
                        {
                            maxValue = value;
                        }
                    }
                }

                return;
            }

            if (chart.IsPercentStacked)
            {
                minValue = 0d;
                maxValue = 1d;
                return;
            }

            var pointCount = 0;
            for (var s = 0; s < chart.Series.Count; s++)
            {
                if (chart.Series[s].Values.Count > pointCount)
                {
                    pointCount = chart.Series[s].Values.Count;
                }
            }

            for (var i = 0; i < pointCount; i++)
            {
                var positive = 0d;
                var negative = 0d;
                var sawValue = false;

                for (var s = 0; s < chart.Series.Count; s++)
                {
                    var series = chart.Series[s];
                    if (i >= series.Values.Count || !series.Values[i].HasValue)
                    {
                        continue;
                    }

                    sawValue = true;
                    var value = series.Values[i].Value;
                    if (value >= 0d)
                    {
                        positive += value;
                    }
                    else
                    {
                        negative += value;
                    }
                }

                if (!sawValue)
                {
                    continue;
                }

                if (negative < minValue)
                {
                    minValue = negative;
                }

                if (positive > maxValue)
                {
                    maxValue = positive;
                }
            }
        }

        public static List<double?> BuildDisplayValues(ParsedChart chart, int seriesIndex)
        {
            var result = new List<double?>();
            if (chart == null || seriesIndex < 0 || seriesIndex >= chart.Series.Count)
            {
                return result;
            }

            var targetSeries = chart.Series[seriesIndex];
            var pointCount = targetSeries.Values.Count;
            if (!chart.IsStacked && !chart.IsPercentStacked)
            {
                for (var i = 0; i < targetSeries.Values.Count; i++)
                {
                    result.Add(targetSeries.Values[i]);
                }

                return result;
            }

            if (chart.IsPercentStacked)
            {
                return BuildPercentStackedDisplayValues(chart, seriesIndex, pointCount);
            }

            for (var i = 0; i < pointCount; i++)
            {
                result.Add(null);
            }

            var positive = new double[pointCount];
            var negative = new double[pointCount];

            for (var s = 0; s <= seriesIndex; s++)
            {
                var series = chart.Series[s];
                var count = Math.Min(pointCount, series.Values.Count);
                for (var i = 0; i < count; i++)
                {
                    if (!series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var value = series.Values[i].Value;
                    if (value >= 0d)
                    {
                        positive[i] += value;
                        if (s == seriesIndex)
                        {
                            result[i] = positive[i];
                        }
                    }
                    else
                    {
                        negative[i] += value;
                        if (s == seriesIndex)
                        {
                            result[i] = negative[i];
                        }
                    }
                }
            }

            return result;
        }

        private static List<double?> BuildPercentStackedDisplayValues(ParsedChart chart, int seriesIndex, int pointCount)
        {
            var result = new List<double?>();
            for (var i = 0; i < pointCount; i++)
            {
                result.Add(null);
            }

            var positiveTotals = new double[pointCount];
            var negativeTotals = new double[pointCount];
            for (var s = 0; s < chart.Series.Count; s++)
            {
                var series = chart.Series[s];
                var count = Math.Min(pointCount, series.Values.Count);
                for (var i = 0; i < count; i++)
                {
                    if (!series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var value = series.Values[i].Value;
                    if (value >= 0d)
                    {
                        positiveTotals[i] += value;
                    }
                    else
                    {
                        negativeTotals[i] += Math.Abs(value);
                    }
                }
            }

            var cumulativePositive = new double[pointCount];
            var cumulativeNegative = new double[pointCount];
            for (var s = 0; s <= seriesIndex; s++)
            {
                var series = chart.Series[s];
                var count = Math.Min(pointCount, series.Values.Count);
                for (var i = 0; i < count; i++)
                {
                    if (!series.Values[i].HasValue)
                    {
                        continue;
                    }

                    var value = series.Values[i].Value;
                    if (value >= 0d)
                    {
                        cumulativePositive[i] += value;
                        if (s == seriesIndex)
                        {
                            result[i] = positiveTotals[i] > 0d ? cumulativePositive[i] / positiveTotals[i] : 0d;
                        }
                    }
                    else
                    {
                        cumulativeNegative[i] += Math.Abs(value);
                        if (s == seriesIndex)
                        {
                            result[i] = negativeTotals[i] > 0d ? -(cumulativeNegative[i] / negativeTotals[i]) : 0d;
                        }
                    }
                }
            }

            return result;
        }
    }
}
