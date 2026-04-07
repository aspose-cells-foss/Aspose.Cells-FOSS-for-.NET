using Aspose.Cells_FOSS;

var options = new LoadOptions
{
    TryRepairPackage = true,
    TryRepairXml = true,
};

try
{
    _ = new Workbook("sample.xlsx", options);
}
catch (WorkbookLoadException exception)
{
    Console.WriteLine(exception.Message);
}
