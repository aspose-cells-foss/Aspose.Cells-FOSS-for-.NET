using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.ListObjects
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "listobjects-sample.xlsx");

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["D1"].PutValue("Quantity");

            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["C2"].PutValue(999.99);
            sheet.Cells["D2"].PutValue(50);

            sheet.Cells["A3"].PutValue("Mouse");
            sheet.Cells["B3"].PutValue("Electronics");
            sheet.Cells["C3"].PutValue(29.99);
            sheet.Cells["D3"].PutValue(200);

            sheet.Cells["A4"].PutValue("Desk Chair");
            sheet.Cells["B4"].PutValue("Furniture");
            sheet.Cells["C4"].PutValue(149.99);
            sheet.Cells["D4"].PutValue(75);

            sheet.Cells["A5"].PutValue("Monitor");
            sheet.Cells["B5"].PutValue("Electronics");
            sheet.Cells["C5"].PutValue(349.99);
            sheet.Cells["D5"].PutValue(100);

            var tableIndex = sheet.ListObjects.Add("A1", "D5", true);
            var table = sheet.ListObjects[tableIndex];
            table.DisplayName = "Products";
            table.Comment = "Product inventory table";
            table.TableStyleType = TableStyleType.TableStyleMedium2;
            table.ShowTableStyleFirstColumn = true;
            table.ShowTableStyleLastColumn = true;
            table.ShowTableStyleRowStripes = true;
            table.ShowTotals = true;

            sheet.Cells["A6"].PutValue("Total");
            sheet.Cells["D6"].PutValue("=SUM(D2:D5)");

            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedSheet = loaded.Worksheets["Data"];
            var loadedTable = loadedSheet.ListObjects["Products"];

            Console.WriteLine("Saved: " + outputPath);
            Console.WriteLine("Table count: " + loadedSheet.ListObjects.Count);
            Console.WriteLine("Table name: " + loadedTable.DisplayName);
            Console.WriteLine("Table comment: " + loadedTable.Comment);
            Console.WriteLine("Table style: " + loadedTable.TableStyleType);
            Console.WriteLine("Show headers: " + loadedTable.ShowHeaderRow);
            Console.WriteLine("Show totals: " + loadedTable.ShowTotals);
            Console.WriteLine("Column count: " + loadedTable.ListColumns.Count);
            Console.WriteLine("First column: " + loadedTable.ListColumns[0].Name);
        }
    }
}
