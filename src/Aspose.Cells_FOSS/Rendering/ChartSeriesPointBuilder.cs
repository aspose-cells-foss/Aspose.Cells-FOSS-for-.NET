using System.Collections.Generic;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartSeriesPointBuilder
    {
        public static List<SKPoint> BuildDateSeriesPoints(ParsedChart chart, List<double?> values, SKRect plot, double min, double max, Core.DateSystem dateSystem)
        {
            if (chart == null || !chart.HasDateCategoryAxis || chart.CategoryValues == null || values == null)
            {
                return null;
            }

            double plotMinDate;
            double plotMaxDate;
            if (!ChartDateAxisMath.TryGetDateAxisPlotRange(chart, dateSystem, out plotMinDate, out plotMaxDate))
            {
                return null;
            }

            var serials = new List<double>();
            var points = new List<SKPoint>();
            var count = chart.CategoryValues.Count;
            if (values.Count < count)
            {
                count = values.Count;
            }

            for (var i = 0; i < count; i++)
            {
                if (!chart.CategoryValues[i].HasValue || !values[i].HasValue)
                {
                    continue;
                }

                var serial = ChartDateAxisMath.NormalizePointSerial(chart, chart.CategoryValues[i].Value, dateSystem);
                serials.Add(serial);
                points.Add(new SKPoint(
                    ChartDateAxisMath.DateCategoryX(serial, plotMinDate, plotMaxDate, plot),
                    ValueToY(values[i].Value, min, max, plot)));
            }

            for (var i = 1; i < serials.Count; i++)
            {
                var serial = serials[i];
                var point = points[i];
                var j = i - 1;
                while (j >= 0 && serials[j] > serial)
                {
                    serials[j + 1] = serials[j];
                    points[j + 1] = points[j];
                    j--;
                }

                serials[j + 1] = serial;
                points[j + 1] = point;
            }

            return points;
        }

        private static float ValueToY(double value, double min, double max, SKRect plot)
        {
            var t = (value - min) / (max - min);
            return plot.Bottom - (float)t * plot.Height;
        }
    }
}
