using Aspose.Cells_FOSS;

var workbook = new Workbook();
var cell = workbook.Worksheets[0].Cells["A1"];
cell.PutValue("Styled");

var style = cell.GetStyle();
style.Font.Bold = true;
style.Pattern = FillPattern.Solid;
style.ForegroundColor = Color.FromArgb(255, 241, 196, 15);
cell.SetStyle(style);

Console.WriteLine($"{cell.StringValue} / Bold={cell.GetStyle().Font.Bold}");
