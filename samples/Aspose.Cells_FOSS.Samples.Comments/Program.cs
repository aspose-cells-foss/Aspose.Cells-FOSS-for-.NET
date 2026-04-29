using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.Comments
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "comments-sample.xlsx");

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Comments";

            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Status");
            sheet.Cells["A2"].PutValue("Design UI");
            sheet.Cells["B2"].PutValue("In Progress");
            sheet.Cells["A3"].PutValue("Implement API");
            sheet.Cells["B3"].PutValue("Pending");
            sheet.Cells["A4"].PutValue("Write Tests");
            sheet.Cells["B4"].PutValue("Completed");

            var comment2 = sheet.Comments.Add(1, 0);
            comment2.Author = "John";
            comment2.Note = "Need to finalize the color scheme by Friday";
            comment2.IsVisible = false;
            comment2.Width = 200;
            comment2.Height = 100;

            var comment3 = sheet.Comments.Add("A3");
            comment3.Author = "Sarah";
            comment3.Note = "Waiting for database schema approval";
            comment3.IsVisible = true;
            comment3.Width = 180;
            comment3.Height = 90;

            var comment4 = sheet.Comments.Add("B4");
            comment4.Author = "Mike";
            comment4.Note = "All unit tests passing. Ready for review.";
            comment4.IsVisible = false;

            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedSheet = loaded.Worksheets["Comments"];

            Console.WriteLine("Saved: " + outputPath);
            Console.WriteLine("Comment count: " + loadedSheet.Comments.Count);
            Console.WriteLine("A2 comment by " + loadedSheet.Comments["A2"].Author + ": " + loadedSheet.Comments["A2"].Note);
            Console.WriteLine("A3 comment visible: " + loadedSheet.Comments["A3"].IsVisible);
            Console.WriteLine("B4 comment: " + loadedSheet.Comments["B4"].Note);
        }
    }
}
