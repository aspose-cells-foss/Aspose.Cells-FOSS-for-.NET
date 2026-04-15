using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.PageSetup
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "page-setup-sample.xlsx");

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Print Sheet";
            sheet.Cells["A1"].PutValue("Title");
            sheet.Cells["C10"].PutValue(42);

            var pageSetup = sheet.PageSetup;
            pageSetup.LeftMarginInch = 0.25d;
            pageSetup.RightMarginInch = 0.4d;
            pageSetup.TopMarginInch = 0.5d;
            pageSetup.BottomMarginInch = 0.6d;
            pageSetup.HeaderMarginInch = 0.2d;
            pageSetup.FooterMarginInch = 0.22d;
            pageSetup.Orientation = PageOrientationType.Landscape;
            pageSetup.PaperSize = PaperSizeType.PaperA4;
            pageSetup.FirstPageNumber = 3;
            pageSetup.Scale = 95;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 2;
            pageSetup.PrintArea = "$A$1:$C$10";
            pageSetup.PrintTitleRows = "$1:$2";
            pageSetup.PrintTitleColumns = "$A:$B";
            pageSetup.LeftHeader = "Left Header";
            pageSetup.CenterHeader = "Center Header";
            pageSetup.RightHeader = "Right Header";
            pageSetup.LeftFooter = "Left Footer";
            pageSetup.CenterFooter = "Center Footer";
            pageSetup.RightFooter = "Right Footer";
            pageSetup.PrintGridlines = true;
            pageSetup.PrintHeadings = true;
            pageSetup.CenterHorizontally = true;
            pageSetup.CenterVertically = true;
            pageSetup.AddHorizontalPageBreak(4);
            pageSetup.AddHorizontalPageBreak(7);
            pageSetup.AddVerticalPageBreak(2);

            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedPageSetup = loaded.Worksheets[0].PageSetup;

            Console.WriteLine("Saved: " + outputPath);
            Console.WriteLine("Orientation: " + loadedPageSetup.Orientation);
            Console.WriteLine("Paper size: " + loadedPageSetup.PaperSize);
            Console.WriteLine("Print area: " + loadedPageSetup.PrintArea);
            Console.WriteLine("Horizontal breaks: " + string.Join(", ", loadedPageSetup.HorizontalPageBreaks));
            Console.WriteLine("Vertical breaks: " + string.Join(", ", loadedPageSetup.VerticalPageBreaks));
        }
    }
}
