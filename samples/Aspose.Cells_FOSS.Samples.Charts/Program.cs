using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.Charts
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "charts-sample.xlsx");

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Charts";

            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            for (var month = 1; month <= 12; month++)
            {
                sheet.Cells[month, 0].PutValue("Month " + month);
                sheet.Cells[month, 1].PutValue(month * 1000 + month * 50);
                sheet.Cells[month, 2].PutValue(month * 300 + month * 20);
            }

            var salesChartIndex = sheet.Charts.Add(ChartType.Column, "Charts!$B$1:$B$13", 0, 4, 18, 8);
            var salesChart = sheet.Charts[salesChartIndex];
            sheet.Cells["E1"].PutValue("Sales Chart: " + salesChart.Name);

            var profitChartIndex = sheet.Charts.Add(ChartType.Line, "Charts!$C$1:$C$13", 0, 9, 18, 13);
            var profitChart = sheet.Charts[profitChartIndex];
            sheet.Cells["J1"].PutValue("Profit Chart: " + profitChart.Name);

            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedSheet = loaded.Worksheets["Charts"];

            Console.WriteLine("Saved: " + outputPath);
            Console.WriteLine("Chart count: " + loadedSheet.Charts.Count);
            Console.WriteLine("First chart: " + loadedSheet.Charts[0].Name + " (Type: " + loadedSheet.Charts[0].ChartType + ")");
            Console.WriteLine("Second chart: " + loadedSheet.Charts[1].Name + " (Type: " + loadedSheet.Charts[1].ChartType + ")");
            Console.WriteLine("First chart anchor: R" + loadedSheet.Charts[0].UpperLeftRow + "C" + loadedSheet.Charts[0].UpperLeftColumn + " to R" + loadedSheet.Charts[0].LowerRightRow + "C" + loadedSheet.Charts[0].LowerRightColumn);
        }
    }
}
