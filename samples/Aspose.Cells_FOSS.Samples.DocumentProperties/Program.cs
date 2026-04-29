using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.DocumentProperties
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "document-properties-sample.xlsx");

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            sheet.Cells["A1"].PutValue("Title");
            sheet.Cells["B1"].PutValue("Document Properties Sample");

            var properties = workbook.DocumentProperties;
            properties.Title = "Annual Sales Report 2024";
            properties.Subject = "Financial Performance Analysis";
            properties.Author = "Finance Department";
            properties.Keywords = "sales, finance, 2024, report";
            properties.Comments = "This document contains quarterly sales data and projections";
            properties.Category = "Financial Reports";
            properties.Company = "Acme Corporation";
            properties.Manager = "Jane Smith";

            properties.Core.Creator = "Finance Department";
            properties.Core.LastModifiedBy = "Finance Department";
            properties.Core.Created = System.DateTime.UtcNow;

            properties.Extended.Company = "Acme Corporation";
            properties.Extended.Manager = "Jane Smith";

            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedProperties = loaded.DocumentProperties;

            Console.WriteLine("Saved: " + outputPath);
            Console.WriteLine("Title: " + loadedProperties.Title);
            Console.WriteLine("Subject: " + loadedProperties.Subject);
            Console.WriteLine("Author: " + loadedProperties.Author);
            Console.WriteLine("Keywords: " + loadedProperties.Keywords);
            Console.WriteLine("Comments: " + loadedProperties.Comments);
            Console.WriteLine("Category: " + loadedProperties.Category);
            Console.WriteLine("Company: " + loadedProperties.Company);
            Console.WriteLine("Manager: " + loadedProperties.Manager);
            Console.WriteLine("Core Creator: " + loadedProperties.Core.Creator);
            Console.WriteLine("Core LastModifiedBy: " + loadedProperties.Core.LastModifiedBy);
            Console.WriteLine("Extended Company: " + loadedProperties.Extended.Company);
            Console.WriteLine("Extended Manager: " + loadedProperties.Extended.Manager);
        }
    }
}
