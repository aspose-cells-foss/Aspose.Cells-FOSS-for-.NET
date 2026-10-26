using System.Collections.Generic;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class MapChartData
    {
        public string Title = string.Empty;
        public string SeriesName = string.Empty;
        public Dictionary<string, double> Values = new Dictionary<string, double>(System.StringComparer.OrdinalIgnoreCase);
        public double Minimum;
        public double Maximum;
        public string MinimumLabel = string.Empty;
        public string MaximumLabel = string.Empty;
        public bool HasCustomColors;
        public bool HasMiddleColor;
        public SKColor MinimumColor;
        public SKColor MiddleColor;
        public SKColor MaximumColor;
    }
}
