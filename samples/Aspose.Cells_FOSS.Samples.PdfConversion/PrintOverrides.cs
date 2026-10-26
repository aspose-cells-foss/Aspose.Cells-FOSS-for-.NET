using System;
using System.Globalization;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.PdfConversion
{
    /// <summary>Explicit output settings for reproducing a print job without editing its XLSX.</summary>
    internal sealed class PrintOverrides
    {
        private PaperSizeType? _paper;
        private int? _scale;
        private readonly double?[] _margins = new double?[4];

        public bool Read(string[] args, ref int index)
        {
            var name = args[index];
            var marginIndex = Array.IndexOf(new[] { "--margin-left-cm", "--margin-right-cm", "--margin-top-cm", "--margin-bottom-cm" }, name);
            if (name != "--paper-size" && name != "--scale" && marginIndex < 0) return false;
            if (index + 1 >= args.Length) throw new ArgumentException(name + " requires a value.");
            var value = args[++index];
            if (name == "--paper-size")
            {
                PaperSizeType paper;
                if (!Enum.TryParse("Paper" + value, true, out paper) || !Enum.IsDefined(typeof(PaperSizeType), paper))
                    throw new ArgumentException("Unsupported paper size: " + value);
                _paper = paper;
            }
            else if (name == "--scale")
            {
                int scale;
                if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out scale) || scale < 10 || scale > 400)
                    throw new ArgumentException("--scale must be between 10 and 400.");
                _scale = scale;
            }
            else
            {
                double margin;
                if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out margin)
                    || double.IsNaN(margin) || double.IsInfinity(margin) || margin < 0d)
                    throw new ArgumentException(name + " requires a non-negative finite number.");
                _margins[marginIndex] = margin;
            }
            return true;
        }

        public void Apply(Workbook workbook)
        {
            for (var index = 0; index < workbook.Worksheets.Count; index++)
            {
                var setup = workbook.Worksheets[index].PageSetup;
                if (_paper.HasValue) setup.PaperSize = _paper.Value;
                if (_scale.HasValue) setup.Scale = _scale.Value;
                if (_margins[0].HasValue) setup.LeftMargin = _margins[0].Value;
                if (_margins[1].HasValue) setup.RightMargin = _margins[1].Value;
                if (_margins[2].HasValue) setup.TopMargin = _margins[2].Value;
                if (_margins[3].HasValue) setup.BottomMargin = _margins[3].Value;
                Console.WriteLine("Print settings: paper=" + setup.PaperSize + ", scale=" + setup.Scale
                    + ", margins (cm)=" + setup.LeftMargin.ToString(CultureInfo.InvariantCulture) + ","
                    + setup.RightMargin.ToString(CultureInfo.InvariantCulture) + ","
                    + setup.TopMargin.ToString(CultureInfo.InvariantCulture) + ","
                    + setup.BottomMargin.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}
