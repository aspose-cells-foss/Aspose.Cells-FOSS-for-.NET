using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.Basic
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "cell-data-roundtrip.xlsx");
            var timestamp = new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc);

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["C1"].PutValue(true);
            sheet.Cells["D1"].PutValue(12.5m);
            sheet.Cells["E1"].PutValue(timestamp);
            sheet.Cells["F1"].PutValue(10);
            sheet.Cells["G1"].PutValue(20);
            sheet.Cells["G1"].Formula = "=F1*2";
            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedSheet = loaded.Worksheets[0];

            Console.WriteLine(loadedSheet.Cells["A1"].StringValue);
            Console.WriteLine(GetValueTypeName(loadedSheet.Cells["B1"]) + ":" + loadedSheet.Cells["B1"].StringValue);
            Console.WriteLine(GetValueTypeName(loadedSheet.Cells["C1"]) + ":" + loadedSheet.Cells["C1"].StringValue);
            Console.WriteLine(GetValueTypeName(loadedSheet.Cells["D1"]) + ":" + loadedSheet.Cells["D1"].StringValue);
            Console.WriteLine(GetValueTypeName(loadedSheet.Cells["E1"]) + ":" + loadedSheet.Cells["E1"].StringValue);
            Console.WriteLine(loadedSheet.Cells["G1"].Formula + " -> " + loadedSheet.Cells["G1"].StringValue);
        }

        private static string GetValueTypeName(Cell cell)
        {
            if (cell.Value == null)
            {
                return string.Empty;
            }

            return cell.Value.GetType().Name;
        }
    }
}
