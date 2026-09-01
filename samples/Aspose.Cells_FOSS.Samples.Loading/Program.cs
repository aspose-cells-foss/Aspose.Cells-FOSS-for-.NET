using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.Loading
{
    internal static class Program
    {
        private static void Main()
        {
            var options = new LoadOptions
            {
                TryRepairPackage = true,
                TryRepairXml = true
            };

            try
            {
                var inputPath = ResolveInputPath();
                new Workbook(inputPath, options);
                Console.WriteLine("Loaded: " + inputPath);
            }
            catch (WorkbookLoadException exception)
            {
                Console.WriteLine(exception.Message);
            }
        }

        private static string ResolveInputPath()
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            var directory = new DirectoryInfo(currentDirectory);

            while (directory != null)
            {
                var inputPath = Path.Combine(directory.FullName, "samples", "files", "sample.xlsx");
                if (File.Exists(inputPath))
                {
                    return inputPath;
                }

                directory = directory.Parent;
            }

            throw new FileNotFoundException("Could not find samples/files/sample.xlsx.");
        }
    }
}
