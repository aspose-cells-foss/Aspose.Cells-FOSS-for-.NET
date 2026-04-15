using System;
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
                new Workbook("sample.xlsx", options);
            }
            catch (WorkbookLoadException exception)
            {
                Console.WriteLine(exception.Message);
            }
        }
    }
}
