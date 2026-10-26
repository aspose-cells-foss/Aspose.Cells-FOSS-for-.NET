using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Word positions for a distributed worksheet line, including its wrap separator.</summary>
    internal sealed class DistributedTextLine
    {
        public string[] Words;
        public float[] Offsets;

        public static float UnwrappedWidth(CellFontContext font, float innerWidth, AlignmentValue alignment)
        {
            var indentStep = Math.Max(font.SizePt * 0.68f, font.Measure("0") * 1.15f);
            var indent = Math.Max(0, alignment.IndentLevel + alignment.RelativeIndent) * indentStep;
            return Math.Max(1f, innerWidth - 2f * indent);
        }

        public static DistributedTextLine Build(CellFontContext font, string line, float width, bool trailingSeparator)
        {
            var words = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            var result = new DistributedTextLine { Words = words, Offsets = new float[words.Length] };
            if (words.Length == 0) return result;
            var widths = new float[words.Length];
            var total = 0f;
            for (var i = 0; i < words.Length; i++)
            {
                widths[i] = font.MeasurePdfText(words[i]);
                total += widths[i];
            }
            if (words.Length == 1)
            {
                result.Offsets[0] = (width - total) / 2f;
                return result;
            }
            // A space that caused an automatic wrap remains a distribution slot at
            // the right edge; a final line or an explicit line break has no such slot.
            var gaps = words.Length - 1 + (trailingSeparator ? 1 : 0);
            var gap = Math.Max(0f, width - total) / gaps;
            for (var i = 1; i < words.Length; i++)
                result.Offsets[i] = result.Offsets[i - 1] + widths[i - 1] + gap;
            return result;
        }
    }
}
