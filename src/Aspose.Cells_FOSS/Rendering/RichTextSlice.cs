using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class RichTextSlice
    {
        public string Text { get; set; }
        public FontValue Font { get; set; }
        public CellFontContext FontContext { get; set; }
        public float WidthPt { get; set; }
        public float AscentPt { get; set; }
        public float DescentPt { get; set; }
        public SKColor Color { get; set; }
    }
}
