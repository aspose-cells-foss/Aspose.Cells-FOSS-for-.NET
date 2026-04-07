using Aspose.Cells_FOSS;

var outputPath = Path.Combine(AppContext.BaseDirectory, "hyperlinks-and-names-sample.xlsx");

var workbook = new Workbook();
var linksSheet = workbook.Worksheets[0];
linksSheet.Name = "Links";
var scopedSheetIndex = workbook.Worksheets.Add("Scoped Sheet");
var scopedSheet = workbook.Worksheets[scopedSheetIndex];
scopedSheet.Cells["B2"].PutValue(5);

linksSheet.Cells["A1"].PutValue("Docs");
var externalLink = linksSheet.Hyperlinks[linksSheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/docs?q=1")];
externalLink.TextToDisplay = "Docs";
externalLink.ScreenTip = "External docs";

linksSheet.Cells["B2"].PutValue("Jump");
var internalLink = linksSheet.Hyperlinks[linksSheet.Hyperlinks.Add("B2", 1, 1, "'Scoped Sheet'!B2")];
internalLink.TextToDisplay = "Jump";
internalLink.ScreenTip = "Jump to scoped sheet";

_ = linksSheet.Hyperlinks[linksSheet.Hyperlinks.Add("C4", "D5", "mailto:test@example.com", "Mail", "Send mail")];

var globalName = workbook.DefinedNames[workbook.DefinedNames.Add("GlobalRange", "='Links'!$A$1:$D$5")];
globalName.Hidden = true;
globalName.Comment = "Primary sample range";

var localName = workbook.DefinedNames[workbook.DefinedNames.Add("ScopedCell", "'Scoped Sheet'!$B$2", scopedSheetIndex)];
localName.Comment = "Sheet-scoped name";

workbook.Save(outputPath);

var loaded = new Workbook(outputPath);
var loadedLinks = loaded.Worksheets["Links"];

Console.WriteLine("Saved: " + outputPath);
Console.WriteLine("Hyperlinks: " + loadedLinks.Hyperlinks.Count);
Console.WriteLine("First hyperlink: " + loadedLinks.Hyperlinks[0].Address + " / " + loadedLinks.Hyperlinks[0].LinkType);
Console.WriteLine("Second hyperlink: " + loadedLinks.Hyperlinks[1].Address + " / " + loadedLinks.Hyperlinks[1].LinkType);
Console.WriteLine("Defined names: " + loaded.DefinedNames.Count);
Console.WriteLine("Global name formula: " + loaded.DefinedNames[0].Formula);
Console.WriteLine("Local name scope: " + (loaded.DefinedNames[1].LocalSheetIndex ?? -1));

