using System;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ScatterChartAxisMath
    {
        public static void ResolveXRange(ParsedChart chart, out double minimum, out double maximum, out double major)
        {
            var dataMinimum = double.MaxValue;
            var dataMaximum = double.MinValue;
            for (var seriesIndex = 0; seriesIndex < chart.Series.Count; seriesIndex++)
            {
                var values = chart.Series[seriesIndex].XValues;
                for (var pointIndex = 0; pointIndex < values.Count; pointIndex++)
                {
                    if (!values[pointIndex].HasValue)
                    {
                        continue;
                    }

                    var value = values[pointIndex].Value;
                    dataMinimum = Math.Min(dataMinimum, value);
                    dataMaximum = Math.Max(dataMaximum, value);
                }
            }

            if (dataMinimum == double.MaxValue)
            {
                minimum = 0d;
                maximum = 1d;
                major = 0.2d;
                return;
            }

            if (chart.IsBubble)
            {
                BubbleChartMath.ResolveXRange(chart, dataMinimum, dataMaximum,
                    out minimum, out maximum, out major);
                return;
            }

            var isDateAxis = ChartXmlParser.IsDateFormat(chart.ScatterXFormatCode);
            var automaticMajor = NiceNumber((dataMaximum - dataMinimum) / 8d);
            var automaticMinimum = !isDateAxis && dataMinimum >= 0d
                ? 0d
                : Math.Floor(dataMinimum / automaticMajor) * automaticMajor;
            minimum = chart.ScatterXMin ?? automaticMinimum;
            major = chart.ScatterXMajorUnit ?? (isDateAxis
                ? automaticMajor
                : NiceNumber((dataMaximum - minimum) / 8d));
            if (major <= 0d)
            {
                major = 1d;
            }

            maximum = chart.ScatterXMax ?? Math.Ceiling(dataMaximum / major) * major;
            if (!chart.ScatterXMax.HasValue && (maximum - dataMaximum) / major < 0.2d)
            {
                maximum += major;
            }
            if (isDateAxis && !chart.ScatterXMin.HasValue && (dataMinimum - minimum) / major < 0.2d)
            {
                minimum -= major;
            }
            if (maximum <= minimum)
            {
                maximum = minimum + major;
            }
        }

        public static float ToScreen(double value, double minimum, double maximum, float left, float right)
        {
            if (maximum <= minimum)
            {
                return left;
            }

            return left + (float)((value - minimum) / (maximum - minimum)) * (right - left);
        }

        private static double NiceNumber(double value)
        {
            if (value <= 0d)
            {
                return 1d;
            }

            var exponent = Math.Floor(Math.Log10(value));
            var fraction = value / Math.Pow(10d, exponent);
            double nice;
            if (fraction < 1.5d) nice = 1d;
            else if (fraction < 3d) nice = 2d;
            else if (fraction < 7d) nice = 5d;
            else nice = 10d;
            return nice * Math.Pow(10d, exponent);
        }
    }
}
