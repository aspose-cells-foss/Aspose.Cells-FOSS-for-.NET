using Aspose.Cells_FOSS;

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
Console.WriteLine(loadedSheet.Cells["B1"].Value?.GetType().Name + ":" + loadedSheet.Cells["B1"].StringValue);
Console.WriteLine(loadedSheet.Cells["C1"].Value?.GetType().Name + ":" + loadedSheet.Cells["C1"].StringValue);
Console.WriteLine(loadedSheet.Cells["D1"].Value?.GetType().Name + ":" + loadedSheet.Cells["D1"].StringValue);
Console.WriteLine(loadedSheet.Cells["E1"].Value?.GetType().Name + ":" + loadedSheet.Cells["E1"].StringValue);
Console.WriteLine(loadedSheet.Cells["G1"].Formula + " -> " + loadedSheet.Cells["G1"].StringValue);
