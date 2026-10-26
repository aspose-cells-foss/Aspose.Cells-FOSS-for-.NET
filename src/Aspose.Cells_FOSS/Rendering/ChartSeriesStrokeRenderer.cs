using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class ChartSeriesStrokeRenderer
    {
        public static void Draw(SKCanvas canvas, SKRect plot, ParsedChart chart, SKPath path, SKPaint paint)
        {
            using (var restore = new SKAutoCanvasRestore(canvas))
            {
                if (chart.Kind == ChartKind.Line && !chart.IsLine3D)
                {
                    // A line on the axis maximum keeps its center at the data value,
                    // but the outer half of its stroke must not escape the plot.
                    canvas.ClipRect(plot);
                }
                canvas.DrawPath(path, paint);
            }
        }
    }
}
