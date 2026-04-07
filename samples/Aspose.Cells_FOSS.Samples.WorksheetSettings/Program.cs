using Aspose.Cells_FOSS;

var outputPath = Path.Combine(AppContext.BaseDirectory, "worksheet-settings-sample.xlsx");

var workbook = new Workbook();
var layout = workbook.Worksheets[0];
layout.Name = "Layout";
layout.VisibilityType = VisibilityType.Hidden;
layout.TabColor = Color.FromArgb(255, 34, 68, 102);
layout.ShowGridlines = false;
layout.ShowRowColumnHeaders = false;
layout.ShowZeros = false;
layout.RightToLeft = true;
layout.Zoom = 85;
layout.Protect();
layout.Protection.Objects = true;
layout.Protection.Scenarios = true;
layout.Protection.FormatCells = true;
layout.Protection.InsertRows = true;
layout.Protection.AutoFilter = true;
layout.Protection.SelectLockedCells = true;
layout.Protection.SelectUnlockedCells = true;

layout.Cells["A1"].PutValue("Merged");
layout.Cells["C4"].PutValue(99);
layout.Cells.Rows[1].Height = 22.5d;
layout.Cells.Rows[3].IsHidden = true;
layout.Cells.Columns[0].Width = 18.25d;
layout.Cells.Columns[2].IsHidden = true;
layout.Cells.Merge(0, 0, 2, 2);

var visibleIndex = workbook.Worksheets.Add("Visible");
var visibleSheet = workbook.Worksheets[visibleIndex];
visibleSheet.Cells["A1"].PutValue("Visible");
workbook.Worksheets.ActiveSheetName = "Visible";

workbook.Save(outputPath);

var loaded = new Workbook(outputPath);
var loadedLayout = loaded.Worksheets["Layout"];

Console.WriteLine("Saved: " + outputPath);
Console.WriteLine("Active sheet: " + loaded.Worksheets.ActiveSheetName);
Console.WriteLine("Layout visibility: " + loadedLayout.VisibilityType);
Console.WriteLine("Merged value: " + loadedLayout.Cells["A1"].StringValue);
Console.WriteLine("Row 2 height: " + (loadedLayout.Cells.Rows[1].Height ?? 0d));
Console.WriteLine("Column A width: " + (loadedLayout.Cells.Columns[0].Width ?? 0d));
Console.WriteLine("Merged regions: " + loadedLayout.Cells.MergedCells.Count);
