using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookHyperlinks;
using static Aspose.Cells_FOSS.XlsxWorkbookDefinedNames;
using static Aspose.Cells_FOSS.XlsxWorkbookPageSetup;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;
using static Aspose.Cells_FOSS.XlsxWorkbookWorksheetWriter;
using static Aspose.Cells_FOSS.XlsxWorkbookWorksheetLoader;
using static Aspose.Cells_FOSS.XlsxWorkbookStyles;
using static Aspose.Cells_FOSS.XlsxWorkbookProperties;
using static Aspose.Cells_FOSS.XlsxDocumentProperties;
using static Aspose.Cells_FOSS.XlsxWorkbookTables;
using static Aspose.Cells_FOSS.XlsxWorkbookPictures;
using static Aspose.Cells_FOSS.XlsxWorkbookComments;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookSerializer
    {
        /// <summary>
        /// Saves the current state.
        /// </summary>
        /// <param name="model">The model.</param>
        /// <param name="stream">The stream.</param>
        /// <param name="options">The options to use.</param>
        public static void Save(WorkbookModel model, Stream stream, SaveOptions options)
        {
            if (options.SaveFormat != SaveFormat.Xlsx)
            {
                throw new UnsupportedFeatureException("Only XLSX save is supported.");
            }

            if (!stream.CanWrite)
            {
                throw new WorkbookSaveException("The output stream must be writable.");
            }

            if (model.Worksheets.Count == 0)
            {
                throw new WorkbookSaveException("A workbook must contain at least one worksheet.");
            }

            var sharedStrings = new SharedStringRepository();
            var stylesheet = BuildStylesheet(model);
            var corePropertiesDocument = BuildCorePropertiesDocument(model);
            var extendedPropertiesDocument = BuildExtendedPropertiesDocument(model);

            foreach (var worksheet in model.Worksheets)
            {
                foreach (var pair in worksheet.Cells)
                {
                    if (!ShouldPersistCell(model.DefaultStyle, pair.Value))
                    {
                        continue;
                    }

                    if (pair.Value.Kind == CellValueKind.String && options.UseSharedStrings)
                    {
                        var text = pair.Value.Value as string;
                        if (text != null)
                        {
                            sharedStrings.Intern(text);
                        }
                    }
                }
            }

            var tableFileOffsets = ComputeTableFileOffsets(model);
            var totalTableCount = ComputeTotalTableCount(model);
            var pictureFileOffsets = ComputePictureFileOffsets(model);
            var chartFileOffsets = ComputeChartFileOffsets(model);
            var drawingNumbers = ComputeDrawingNumbers(model);
            var totalDrawingCount = ComputeTotalDrawingCount(model);
            var imageExtensions = CollectImageExtensions(model);
            var chartPartNames = CollectChartPartNames(model, chartFileOffsets);
            var chartContentTypes = CollectChartContentTypes(model);
            var chartCompanionPartNames = CollectChartCompanionPartNames(model, chartFileOffsets, totalDrawingCount);
            var chartCompanionContentTypes = CollectChartCompanionContentTypes(model);
            var commentFileNumbers = ComputeCommentFileNumbers(model);
            var totalCommentCount = ComputeTotalCommentCount(model);

            var userShapesDrawingCounter = totalDrawingCount;

            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
                var hasSharedStrings = sharedStrings.Values.Count > 0;
                var hasTheme = !string.IsNullOrEmpty(model.RawThemeXml);
                var externalLinkBaseRId = model.Worksheets.Count
                    + (hasSharedStrings ? 1 : 0)
                    + 1
                    + (hasTheme ? 1 : 0)
                    + 1;
                WriteXmlEntry(archive, "[Content_Types].xml", BuildContentTypes(model, hasSharedStrings, true, corePropertiesDocument != null, extendedPropertiesDocument != null, totalTableCount, totalDrawingCount, imageExtensions, chartPartNames, chartContentTypes, chartCompanionPartNames, chartCompanionContentTypes, hasTheme, totalCommentCount));
                WriteXmlEntry(archive, "_rels/.rels", BuildRootRelationships(corePropertiesDocument != null, extendedPropertiesDocument != null));
                WriteXmlEntry(archive, "xl/workbook.xml", BuildWorkbook(model, externalLinkBaseRId));
                WriteXmlEntry(archive, "xl/_rels/workbook.xml.rels", BuildWorkbookRelationships(model, hasSharedStrings, true));

                if (hasTheme)
                {
                    var themeEntry = archive.CreateEntry("xl/theme/theme1.xml", CompressionLevel.Optimal);
                    using (var themeStream = themeEntry.Open())
                    using (var writer = new StreamWriter(themeStream, new UTF8Encoding(false)))
                    {
                        writer.Write(model.RawThemeXml);
                    }
                }

                for (var e = 0; e < model.ExternalLinks.Count; e++)
                {
                    var extLink = model.ExternalLinks[e];
                    var extN = (e + 1).ToString(CultureInfo.InvariantCulture);
                    var extPath = "xl/externalLinks/externalLink" + extN + ".xml";
                    var extEntry = archive.CreateEntry(extPath, CompressionLevel.Optimal);
                    using (var extStream = extEntry.Open())
                    using (var writer = new StreamWriter(extStream, new UTF8Encoding(false)))
                    {
                        writer.Write(extLink.RawXml);
                    }

                    if (!string.IsNullOrEmpty(extLink.RawRelsXml))
                    {
                        var extRelsPath = "xl/externalLinks/_rels/externalLink" + extN + ".xml.rels";
                        var extRelsEntry = archive.CreateEntry(extRelsPath, CompressionLevel.Optimal);
                        using (var extRelsStream = extRelsEntry.Open())
                        using (var writer = new StreamWriter(extRelsStream, new UTF8Encoding(false)))
                        {
                            writer.Write(extLink.RawRelsXml);
                        }
                    }
                }

                for (var i = 0; i < model.Worksheets.Count; i++)
                {
                    var worksheet = model.Worksheets[i];
                    var tableFileOffset = tableFileOffsets[i];
                    var pictureFileOffset = pictureFileOffsets[i];
                    var chartFileOffset = chartFileOffsets[i];
                    var drawingNumber = drawingNumbers[i];
                    var commentFileNumber = commentFileNumbers[i];
                    var hasDrawing = worksheet.Pictures.Count > 0 || worksheet.Shapes.Count > 0 || worksheet.Charts.Count > 0;
                    var hasComments = worksheet.Comments.Count > 0;
                    var externalHyperlinkCount = CountExternalHyperlinks(worksheet);

                    for (var t = 0; t < worksheet.ListObjects.Count; t++)
                    {
                        var globalTableNumber = tableFileOffset + t + 1;
                        var tableDocument = BuildTableDocument(worksheet.ListObjects[t], globalTableNumber);
                        WriteXmlEntry(archive, "xl/tables/table" + globalTableNumber + ".xml", tableDocument);
                    }

                    if (hasDrawing)
                    {
                        WriteXmlEntry(archive, "xl/drawings/drawing" + drawingNumber + ".xml", BuildDrawingDocument(worksheet, pictureFileOffset, chartFileOffset));
                        if (worksheet.Pictures.Count > 0 || worksheet.ShapeImages.Count > 0)
                        {
                            WritePictureMediaEntries(archive, worksheet, pictureFileOffset);
                        }

                        if (worksheet.Pictures.Count > 0 || worksheet.ShapeImages.Count > 0 || worksheet.Charts.Count > 0)
                        {
                            WriteXmlEntry(archive, "xl/drawings/_rels/drawing" + drawingNumber + ".xml.rels", BuildDrawingRelationshipsDocument(worksheet, pictureFileOffset, chartFileOffset));
                        }

                        for (var k = 0; k < worksheet.Charts.Count; k++)
                        {
                            var globalChartNumber = chartFileOffset + k + 1;
                            WriteChartFiles(archive, worksheet.Charts[k], globalChartNumber, ref userShapesDrawingCounter);
                        }
                    }

                    if (hasComments)
                    {
                        WriteXmlEntry(archive, "xl/comments" + commentFileNumber + ".xml", BuildCommentsDocument(worksheet));
                        WriteVmlDrawing(archive, worksheet, commentFileNumber);
                    }

                    WriteXmlEntry(archive, "xl/worksheets/sheet" + (i + 1) + ".xml", BuildWorksheet(worksheet, model.DefaultStyle, model.Settings.DateSystem, sharedStrings, options, stylesheet, externalHyperlinkCount, hasDrawing, hasComments));

                    var worksheetRelationships = BuildWorksheetRelationshipsDocument(worksheet, tableFileOffset, drawingNumber, commentFileNumber);
                    if (worksheetRelationships != null)
                    {
                        WriteXmlEntry(archive, "xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels", worksheetRelationships);
                    }
                }

                if (sharedStrings.Values.Count > 0)
                {
                    WriteXmlEntry(archive, "xl/sharedStrings.xml", BuildSharedStrings(sharedStrings));
                }

                WriteXmlEntry(archive, "xl/styles.xml", stylesheet.Document);

                if (corePropertiesDocument != null)
                {
                    WriteXmlEntry(archive, "docProps/core.xml", corePropertiesDocument);
                }

                if (extendedPropertiesDocument != null)
                {
                    WriteXmlEntry(archive, "docProps/app.xml", extendedPropertiesDocument);
                }
            }
        }

        private static int[] ComputeTableFileOffsets(WorkbookModel model)
        {
            var offsets = new int[model.Worksheets.Count];
            var running = 0;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                offsets[i] = running;
                running += model.Worksheets[i].ListObjects.Count;
            }

            return offsets;
        }

        private static int ComputeTotalTableCount(WorkbookModel model)
        {
            var total = 0;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                total += model.Worksheets[i].ListObjects.Count;
            }

            return total;
        }

        private static int CountExternalHyperlinks(WorksheetModel worksheet)
        {
            var count = 0;
            var ordered = GetOrderedHyperlinks(worksheet.Hyperlinks);
            for (var i = 0; i < ordered.Count; i++)
            {
                if (!string.IsNullOrEmpty(ordered[i].Address))
                {
                    count++;
                }
            }

            return count;
        }

        private static int[] ComputePictureFileOffsets(WorkbookModel model)
        {
            var offsets = new int[model.Worksheets.Count];
            var running = 0;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                offsets[i] = running;
                running += model.Worksheets[i].Pictures.Count + model.Worksheets[i].ShapeImages.Count;
            }

            return offsets;
        }

        private static int[] ComputeChartFileOffsets(WorkbookModel model)
        {
            var offsets = new int[model.Worksheets.Count];
            var running = 0;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                offsets[i] = running;
                running += model.Worksheets[i].Charts.Count;
            }

            return offsets;
        }

        private static int[] ComputeDrawingNumbers(WorkbookModel model)
        {
            var numbers = new int[model.Worksheets.Count];
            var drawingCounter = 0;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                var ws = model.Worksheets[i];
                if (ws.Pictures.Count > 0 || ws.Shapes.Count > 0 || ws.Charts.Count > 0)
                {
                    drawingCounter++;
                    numbers[i] = drawingCounter;
                }
                else
                {
                    numbers[i] = 0;
                }
            }

            return numbers;
        }

        private static int[] ComputeCommentFileNumbers(WorkbookModel model)
        {
            var numbers = new int[model.Worksheets.Count];
            var counter = 0;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                if (model.Worksheets[i].Comments.Count > 0)
                {
                    counter++;
                    numbers[i] = counter;
                }
                else
                {
                    numbers[i] = 0;
                }
            }

            return numbers;
        }

        private static int ComputeTotalCommentCount(WorkbookModel model)
        {
            var total = 0;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                if (model.Worksheets[i].Comments.Count > 0)
                {
                    total++;
                }
            }

            return total;
        }

        private static int ComputeTotalDrawingCount(WorkbookModel model)
        {
            var total = 0;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                var ws = model.Worksheets[i];
                if (ws.Pictures.Count > 0 || ws.Shapes.Count > 0 || ws.Charts.Count > 0)
                {
                    total++;
                }
            }

            return total;
        }

        private static IReadOnlyList<string> CollectChartCompanionPartNames(WorkbookModel model, int[] chartFileOffsets, int totalDrawingCount)
        {
            var result = new List<string>();
            var userShapesCounter = totalDrawingCount;
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                var charts = model.Worksheets[i].Charts;
                for (var k = 0; k < charts.Count; k++)
                {
                    var globalChartNumber = chartFileOffsets[i] + k + 1;
                    foreach (var companion in charts[k].CompanionFiles)
                    {
                        string partName;
                        if (companion.FileName.IndexOf('/') >= 0)
                        {
                            userShapesCounter++;
                            partName = "/xl/drawings/drawing" + userShapesCounter.ToString(CultureInfo.InvariantCulture) + ".xml";
                        }
                        else
                        {
                            partName = "/" + ResolveCompanionPath(RenumberCompanionFileName(companion.FileName, globalChartNumber));
                        }

                        result.Add(partName);
                    }
                }
            }

            return result;
        }

        private static IReadOnlyList<string> CollectChartCompanionContentTypes(WorkbookModel model)
        {
            var result = new List<string>();
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                var charts = model.Worksheets[i].Charts;
                for (var k = 0; k < charts.Count; k++)
                {
                    foreach (var companion in charts[k].CompanionFiles)
                    {
                        result.Add(GetCompanionContentType(companion.RelationshipType));
                    }
                }
            }

            return result;
        }

        private static IReadOnlyList<string> CollectChartPartNames(WorkbookModel model, int[] chartFileOffsets)
        {
            var result = new List<string>();
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                var charts = model.Worksheets[i].Charts;
                for (var k = 0; k < charts.Count; k++)
                {
                    var globalChartNumber = (chartFileOffsets[i] + k + 1).ToString(CultureInfo.InvariantCulture);
                    var fileName = charts[k].IsChartEx
                        ? "/xl/charts/chartEx" + globalChartNumber + ".xml"
                        : "/xl/charts/chart" + globalChartNumber + ".xml";
                    result.Add(fileName);
                }
            }

            return result;
        }

        private static IReadOnlyList<string> CollectChartContentTypes(WorkbookModel model)
        {
            var result = new List<string>();
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                var charts = model.Worksheets[i].Charts;
                for (var k = 0; k < charts.Count; k++)
                {
                    result.Add(charts[k].IsChartEx ? ChartExContentType : ChartContentType);
                }
            }

            return result;
        }

        private static string GetCompanionContentType(string relationshipType)
        {
            if (string.Equals(relationshipType, ChartStyleRelationshipType, StringComparison.OrdinalIgnoreCase))
            {
                return ChartStyleContentType;
            }

            if (string.Equals(relationshipType, ChartColorStyleRelationshipType, StringComparison.OrdinalIgnoreCase))
            {
                return ChartColorStyleContentType;
            }

            if (string.Equals(relationshipType, ChartUserShapesRelationshipType, StringComparison.OrdinalIgnoreCase))
            {
                return ChartUserShapesContentType;
            }

            return "application/octet-stream";
        }

        private static string RenumberCompanionFileName(string originalName, int globalNumber)
        {
            // Cross-directory paths (e.g. "../drawings/drawing4.xml") are kept verbatim
            // to avoid overwriting worksheet drawings with chart user-shape files.
            if (originalName.IndexOf('/') >= 0)
            {
                return originalName;
            }

            var suffixStart = originalName.LastIndexOf('.');
            if (suffixStart < 0)
            {
                return originalName;
            }

            var suffix = originalName.Substring(suffixStart);
            var basePart = originalName.Substring(0, suffixStart);
            var trimmed = basePart.TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');
            if (string.IsNullOrEmpty(trimmed) || trimmed == basePart)
            {
                return originalName;
            }

            return trimmed + globalNumber.ToString(CultureInfo.InvariantCulture) + suffix;
        }

        // Resolves a chart companion target relative to "xl/charts/" and normalises ".." segments.
        // e.g. ("../drawings/drawing4.xml") → "xl/drawings/drawing4.xml"
        //      ("style1.xml")              → "xl/charts/style1.xml"
        private static string ResolveCompanionPath(string renumberedTarget)
        {
            var parts = ("xl/charts/" + renumberedTarget).Split('/');
            var stack = new List<string>();
            foreach (var part in parts)
            {
                if (part == "..")
                {
                    if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                }
                else if (part.Length > 0 && part != ".")
                {
                    stack.Add(part);
                }
            }

            return string.Join("/", stack);
        }

        private static void WriteChartFiles(ZipArchive archive, ChartModel chart, int globalChartNumber, ref int userShapesDrawingCounter)
        {
            var chartFileName = chart.IsChartEx
                ? "chartEx" + globalChartNumber.ToString(CultureInfo.InvariantCulture) + ".xml"
                : "chart" + globalChartNumber.ToString(CultureInfo.InvariantCulture) + ".xml";
            var chartPath = "xl/charts/" + chartFileName;
            var chartEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
            using (var stream = chartEntry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(chart.RawChartXml);
            }

            if (chart.CompanionFiles.Count == 0)
            {
                return;
            }

            var companionRels = new XElement(PackageRelationshipNs + "Relationships");
            foreach (var companion in chart.CompanionFiles)
            {
                string companionPath;
                string relsTarget;

                if (companion.FileName.IndexOf('/') >= 0)
                {
                    // Cross-directory companion (e.g. chart user shapes): assign the next
                    // drawing number after all worksheet drawings to avoid collisions.
                    userShapesDrawingCounter++;
                    var drawingName = "drawing" + userShapesDrawingCounter.ToString(CultureInfo.InvariantCulture) + ".xml";
                    companionPath = "xl/drawings/" + drawingName;
                    relsTarget = "../drawings/" + drawingName;
                }
                else
                {
                    var newFileName = RenumberCompanionFileName(companion.FileName, globalChartNumber);
                    companionPath = ResolveCompanionPath(newFileName);
                    relsTarget = newFileName;
                }

                var companionEntry = archive.CreateEntry(companionPath, CompressionLevel.Optimal);
                using (var stream = companionEntry.Open())
                using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                {
                    writer.Write(companion.RawContent);
                }

                companionRels.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", companion.RelationshipId),
                    new XAttribute("Type", companion.RelationshipType),
                    new XAttribute("Target", relsTarget)));
            }

            var relsPath = "xl/charts/_rels/" + chartFileName + ".rels";
            WriteXmlEntry(archive, relsPath, new XDocument(new XDeclaration("1.0", "utf-8", "yes"), companionRels));
        }

        private static IReadOnlyList<string> CollectImageExtensions(WorkbookModel model)
        {
            var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<string>();
            for (var i = 0; i < model.Worksheets.Count; i++)
            {
                var ws = model.Worksheets[i];
                for (var p = 0; p < ws.Pictures.Count; p++)
                {
                    var ext = ws.Pictures[p].ImageExtension;
                    if (!string.IsNullOrEmpty(ext) && seen.Add(ext))
                    {
                        result.Add(ext);
                    }
                }
                for (var s = 0; s < ws.ShapeImages.Count; s++)
                {
                    var ext = ws.ShapeImages[s].Extension;
                    if (!string.IsNullOrEmpty(ext) && seen.Add(ext))
                    {
                        result.Add(ext);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Loads the current state.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="options">The options to use.</param>
        /// <param name="diagnostics">The diagnostics.</param>
        /// <returns>The workbook model.</returns>
        public static WorkbookModel Load(Stream stream, LoadOptions options, LoadDiagnostics diagnostics)
        {
            if (!stream.CanRead)
            {
                throw new WorkbookLoadException("The input stream must be readable.");
            }

            var workingStream = EnsureSeekable(stream);

            try
            {
                using (var archive = new ZipArchive(workingStream, ZipArchiveMode.Read, true))
                {
                    var workbookEntry = GetEntry(archive, "/xl/workbook.xml");
                    if (workbookEntry == null)
                    {
                        throw new InvalidFileFormatException("The package does not contain /xl/workbook.xml.");
                    }

                    var workbookDocument = LoadDocument(workbookEntry);
                    var workbookRelationships = LoadRelationships(archive, "/xl/_rels/workbook.xml.rels", "/xl/workbook.xml");
                    var sharedStrings = LoadSharedStrings(archive, workbookRelationships, options, diagnostics);
                    var stylesheet = LoadStylesheet(archive, workbookRelationships, options, diagnostics);

                    var workbookModel = new WorkbookModel();
                    workbookModel.Worksheets.Clear();
                    workbookModel.DefaultStyle = stylesheet.DefaultCellStyle.Clone();

                    var workbookRoot = workbookDocument.Root;
                    if (workbookRoot == null)
                    {
                        throw new InvalidFileFormatException("The workbook XML is empty.");
                    }

                    var worksheetDefinedNames = LoadWorksheetDefinedNames(workbookRoot, diagnostics, options);

                    var sheets = workbookRoot.Element(MainNs + "sheets");
                    if (sheets == null)
                    {
                        throw new InvalidFileFormatException("The workbook XML does not contain a sheets element.");
                    }

                    var sheetElements = new List<XElement>(sheets.Elements(MainNs + "sheet"));
                    if (sheetElements.Count == 0)
                    {
                        throw new InvalidFileFormatException("The workbook XML does not contain any worksheets.");
                    }

                    LoadWorkbookMetadata(workbookRoot, workbookModel, sheetElements.Count, diagnostics, options);
                    LoadDocumentProperties(archive, workbookModel, diagnostics, options);
                    LoadWorkbookDefinedNames(workbookRoot, workbookModel, sheetElements.Count, diagnostics, options);
                    workbookModel.RawThemeXml = LoadRawTheme(archive);
                    LoadExternalLinks(archive, workbookRoot, workbookRelationships, workbookModel);

                    workbookModel.ActiveSheetIndex = workbookModel.ActiveSheetIndex < sheetElements.Count ? workbookModel.ActiveSheetIndex : 0;
                    for (var index = 0; index < sheetElements.Count; index++)
                    {
                        var sheetElement = sheetElements[index];
                        var sheetName = (string)sheetElement.Attribute("name");
                        if (string.IsNullOrWhiteSpace(sheetName))
                        {
                            sheetName = "Sheet" + (index + 1);
                        }

                        var resolvedSheetName = sheetName;

                        var relationshipId = (string)sheetElement.Attribute(RelationshipNs + "id");
                        var worksheetUri = ResolveWorksheetUri(index, relationshipId, workbookRelationships, archive, diagnostics, resolvedSheetName, options);
                        var worksheetEntry = GetEntry(archive, worksheetUri);
                        if (worksheetEntry == null)
                        {
                            throw new InvalidFileFormatException("Worksheet part '" + worksheetUri + "' was not found.");
                        }

                        var worksheetHyperlinkTargets = LoadWorksheetHyperlinkTargets(archive, worksheetUri);

                        WorksheetDefinedNamesState definedNamesState;
                        worksheetDefinedNames.TryGetValue(index, out definedNamesState);
                        var worksheetModel = LoadWorksheet(
                            resolvedSheetName,
                            LoadDocument(worksheetEntry),
                            worksheetHyperlinkTargets,
                            workbookModel.Settings.DateSystem,
                            sharedStrings,
                            stylesheet,
                            diagnostics,
                            options,
                            definedNamesState,
                            archive,
                            worksheetUri);

                        var state = (string)sheetElement.Attribute("state");
                        switch (state)
                        {
                            case "hidden":
                                worksheetModel.Visibility = SheetVisibility.Hidden;
                                break;
                            case "veryHidden":
                                worksheetModel.Visibility = SheetVisibility.VeryHidden;
                                break;
                            default:
                                worksheetModel.Visibility = SheetVisibility.Visible;
                                break;
                        }

                        workbookModel.Worksheets.Add(worksheetModel);
                    }

                    return workbookModel;
                }
            }
            catch (CellsException)
            {
                throw;
            }
            catch (InvalidDataException exception)
            {
                throw new InvalidFileFormatException("The workbook is not a valid XLSX zip package.", exception);
            }
            finally
            {
                if (!ReferenceEquals(workingStream, stream))
                {
                    workingStream.Dispose();
                }
            }
        }

        private static string LoadRawTheme(ZipArchive archive)
        {
            var entry = GetEntry(archive, "xl/theme/theme1.xml");
            if (entry == null)
            {
                return null;
            }

            using (var stream = entry.Open())
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private static void LoadExternalLinks(ZipArchive archive, XElement workbookRoot, IReadOnlyDictionary<string, string> workbookRelationships, WorkbookModel workbookModel)
        {
            var externalReferencesElement = workbookRoot.Element(MainNs + "externalReferences");
            if (externalReferencesElement == null)
            {
                return;
            }

            foreach (var extRef in externalReferencesElement.Elements(MainNs + "externalReference"))
            {
                var rId = (string)extRef.Attribute(RelationshipNs + "id");
                if (string.IsNullOrEmpty(rId))
                {
                    continue;
                }

                string partUri;
                if (!workbookRelationships.TryGetValue(rId, out partUri))
                {
                    continue;
                }

                var partEntry = GetEntry(archive, partUri);
                if (partEntry == null)
                {
                    continue;
                }

                string rawXml;
                using (var rawStream = partEntry.Open())
                using (var reader = new StreamReader(rawStream, Encoding.UTF8))
                {
                    rawXml = reader.ReadToEnd();
                }

                var normalizedUri = partUri.TrimStart('/');
                var lastSlash = normalizedUri.LastIndexOf('/');
                var directory = lastSlash >= 0 ? normalizedUri.Substring(0, lastSlash + 1) : string.Empty;
                var fileName = lastSlash >= 0 ? normalizedUri.Substring(lastSlash + 1) : normalizedUri;
                var relsUri = "/" + directory + "_rels/" + fileName + ".rels";

                string rawRelsXml = null;
                var relsEntry = GetEntry(archive, relsUri);
                if (relsEntry != null)
                {
                    using (var relsStream = relsEntry.Open())
                    using (var reader = new StreamReader(relsStream, Encoding.UTF8))
                    {
                        rawRelsXml = reader.ReadToEnd();
                    }
                }

                workbookModel.ExternalLinks.Add(new ExternalLinkModel { RawXml = rawXml, RawRelsXml = rawRelsXml });
            }
        }
    }
}
