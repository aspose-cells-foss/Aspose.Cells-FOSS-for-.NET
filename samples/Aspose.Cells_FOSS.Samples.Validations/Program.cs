using Aspose.Cells_FOSS;

var outputPath = Path.Combine(AppContext.BaseDirectory, "validations-sample.xlsx");

var workbook = new Workbook();
var sheet = workbook.Worksheets[0];
sheet.Name = "Validation Sheet";

sheet.Cells["A1"].PutValue("Open");
sheet.Cells["B2"].PutValue(5);
sheet.Cells["C3"].PutValue(7);
sheet.Cells["E2"].PutValue(8);
sheet.Cells["G1"].PutValue("ABCDE");

var listValidationIndex = sheet.Validations.Add(CellArea.CreateCellArea("A1", "A3"));
var listValidation = sheet.Validations[listValidationIndex];
listValidation.Type = ValidationType.List;
listValidation.Formula1 = "\"Open,Closed\"";
listValidation.IgnoreBlank = true;
listValidation.InCellDropDown = true;
listValidation.ShowInput = true;
listValidation.InputTitle = "Status";
listValidation.InputMessage = "Pick a status";
listValidation.ShowError = true;
listValidation.ErrorTitle = "Invalid";
listValidation.ErrorMessage = "Choose from the list";

var decimalValidationIndex = sheet.Validations.Add(CellArea.CreateCellArea("B2", "C3"));
var decimalValidation = sheet.Validations[decimalValidationIndex];
decimalValidation.Type = ValidationType.Decimal;
decimalValidation.Operator = OperatorType.Between;
decimalValidation.Formula1 = "1.5";
decimalValidation.Formula2 = "9.5";
decimalValidation.ErrorTitle = "Range";
decimalValidation.ErrorMessage = "Enter 1.5-9.5";
decimalValidation.ShowError = true;
decimalValidation.AddArea(CellArea.CreateCellArea("E2", "E3"));

var customValidationIndex = sheet.Validations.Add(CellArea.CreateCellArea("G1", "G1"));
var customValidation = sheet.Validations[customValidationIndex];
customValidation.Type = ValidationType.Custom;
customValidation.AlertStyle = ValidationAlertType.Warning;
customValidation.Formula1 = "LEN(G1)<=5";
customValidation.ShowInput = true;
customValidation.InputTitle = "Code";
customValidation.InputMessage = "Up to 5 chars";

workbook.Save(outputPath);

var loaded = new Workbook(outputPath);
var loadedSheet = loaded.Worksheets["Validation Sheet"];
var a1Validation = loadedSheet.Validations.GetValidationInCell(0, 0);
var c2Validation = loadedSheet.Validations.GetValidationInCell(1, 2);

Console.WriteLine("Saved: " + outputPath);
Console.WriteLine("Validation count: " + loadedSheet.Validations.Count);
Console.WriteLine("A1 validation: " + a1Validation!.Type + " / " + a1Validation.Formula1);
Console.WriteLine("C2 validation: " + c2Validation!.Type + " / " + c2Validation.Formula1 + " to " + c2Validation.Formula2);
Console.WriteLine("G1 validation: " + loadedSheet.Validations.GetValidationInCell(0, 6)!.Type);
