using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookPictures;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookPicturesLoader
    {
        internal static void LoadPictures(WorksheetModel worksheetModel, XElement worksheetRoot, ZipArchive archive, string worksheetUri, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            worksheetModel.Pictures.Clear();
            worksheetModel.ShapeImages.Clear();
            worksheetModel.Shapes.Clear();
            worksheetModel.Charts.Clear();

            var drawingUri = FindDrawingUri(archive, worksheetUri);
            if (string.IsNullOrEmpty(drawingUri))
            {
                return;
            }

            var drawingEntry = GetEntry(archive, drawingUri);
            if (drawingEntry == null)
            {
                AddIssue(diagnostics, options, new LoadIssue("PIC-R001", DiagnosticSeverity.Recoverable, "Drawing part '" + drawingUri + "' was referenced but not found; pictures were skipped.", repairApplied: true)
                {
                    SheetName = sheetName,
                });
                return;
            }

            var drawingDocument = LoadDocument(drawingEntry);
            var drawingRoot = drawingDocument.Root;
            if (drawingRoot == null)
            {
                return;
            }

            var imageTargets = LoadDrawingImageTargets(archive, drawingUri);
            var chartTargets = LoadDrawingChartTargets(archive, drawingUri);

            LoadTwoCellAnchorPictures(worksheetModel, drawingRoot, imageTargets, archive, diagnostics, options, sheetName);
            LoadOneCellAnchorPictures(worksheetModel, drawingRoot, imageTargets, archive, diagnostics, options, sheetName);
            LoadShapeImages(worksheetModel, imageTargets, archive);
            LoadTwoCellAnchorShapes(worksheetModel, drawingRoot);
            LoadOneCellAnchorShapes(worksheetModel, drawingRoot);
            LoadTwoCellAnchorCharts(worksheetModel, drawingRoot, chartTargets, archive, diagnostics, options, sheetName);
            LoadOneCellAnchorCharts(worksheetModel, drawingRoot, chartTargets, archive, diagnostics, options, sheetName);
        }

        private static void LoadTwoCellAnchorPictures(WorksheetModel worksheetModel, XElement drawingRoot, IReadOnlyDictionary<string, string> imageTargets, ZipArchive archive, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            foreach (var anchor in drawingRoot.Elements(XdrNs + "twoCellAnchor"))
            {
                var pic = anchor.Element(XdrNs + "pic");
                if (pic == null)
                {
                    continue;
                }

                var fromEl = anchor.Element(XdrNs + "from");
                var toEl = anchor.Element(XdrNs + "to");
                if (fromEl == null || toEl == null)
                {
                    continue;
                }

                var model = ParsePicture(pic, fromEl, toEl, imageTargets, archive, diagnostics, options, sheetName);
                if (model != null)
                {
                    worksheetModel.Pictures.Add(model);
                }
            }
        }

        private static void LoadOneCellAnchorPictures(WorksheetModel worksheetModel, XElement drawingRoot, IReadOnlyDictionary<string, string> imageTargets, ZipArchive archive, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            foreach (var anchor in drawingRoot.Elements(XdrNs + "oneCellAnchor"))
            {
                var pic = anchor.Element(XdrNs + "pic");
                if (pic == null)
                {
                    continue;
                }

                var fromEl = anchor.Element(XdrNs + "from");
                if (fromEl == null)
                {
                    continue;
                }

                var extEl = anchor.Element(XdrNs + "ext");
                var model = ParseOneCellPicture(pic, fromEl, extEl, imageTargets, archive, diagnostics, options, sheetName);
                if (model != null)
                {
                    worksheetModel.Pictures.Add(model);
                }
            }
        }

        private static void LoadShapeImages(WorksheetModel worksheetModel, IReadOnlyDictionary<string, string> imageTargets, ZipArchive archive)
        {
            worksheetModel.ShapeImages.Clear();

            // Collect rIds already claimed by loaded top-level pictures
            var claimedRIds = new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);
            foreach (var pic in worksheetModel.Pictures)
            {
                if (!string.IsNullOrEmpty(pic.OriginalRId))
                {
                    claimedRIds.Add(pic.OriginalRId);
                }
            }

            // Load every image relationship that was NOT claimed by a top-level picture.
            // These images are referenced from within shape/group raw XML and must be preserved
            // so their r:embed rId references remain valid after save.
            foreach (var kvp in imageTargets)
            {
                var rId = kvp.Key;
                if (claimedRIds.Contains(rId))
                {
                    continue;
                }

                var entry = GetEntry(archive, kvp.Value);
                if (entry == null)
                {
                    continue;
                }

                byte[] imageData;
                using (var stream = entry.Open())
                using (var ms = new MemoryStream())
                {
                    stream.CopyTo(ms);
                    imageData = ms.ToArray();
                }

                if (imageData.Length == 0)
                {
                    continue;
                }

                worksheetModel.ShapeImages.Add(new Core.ShapeImageModel
                {
                    OriginalRId = rId,
                    Extension = Picture.DetectExtension(imageData),
                    ImageData = imageData,
                });
            }
        }

        private static void LoadTwoCellAnchorShapes(WorksheetModel worksheetModel, XElement drawingRoot)
        {
            foreach (var anchor in drawingRoot.Elements(XdrNs + "twoCellAnchor"))
            {
                var fromEl = anchor.Element(XdrNs + "from");
                var toEl = anchor.Element(XdrNs + "to");
                if (fromEl == null || toEl == null)
                {
                    continue;
                }

                var sp = anchor.Element(XdrNs + "sp");
                if (sp != null)
                {
                    var model = ParseShape(sp, fromEl, toEl);
                    if (model != null)
                    {
                        worksheetModel.Shapes.Add(model);
                    }
                    continue;
                }

                var grpSp = anchor.Element(XdrNs + "grpSp");
                if (grpSp != null)
                {
                    var model = ParseGroupShape(grpSp, fromEl, toEl);
                    if (model != null)
                    {
                        worksheetModel.Shapes.Add(model);
                    }
                    continue;
                }

                var cxnSp = anchor.Element(XdrNs + "cxnSp");
                if (cxnSp != null)
                {
                    var model = ParseConnector(cxnSp, fromEl, toEl);
                    if (model != null)
                    {
                        worksheetModel.Shapes.Add(model);
                    }
                }
            }
        }

        private static void LoadOneCellAnchorShapes(WorksheetModel worksheetModel, XElement drawingRoot)
        {
            foreach (var anchor in drawingRoot.Elements(XdrNs + "oneCellAnchor"))
            {
                var fromEl = anchor.Element(XdrNs + "from");
                if (fromEl == null)
                {
                    continue;
                }

                var extEl = anchor.Element(XdrNs + "ext");

                var sp = anchor.Element(XdrNs + "sp");
                if (sp != null)
                {
                    var model = ParseOneCellShape(sp, fromEl, extEl);
                    if (model != null)
                    {
                        worksheetModel.Shapes.Add(model);
                    }
                    continue;
                }

                var grpSp = anchor.Element(XdrNs + "grpSp");
                if (grpSp != null)
                {
                    var model = ParseOneCellGroupShape(grpSp, fromEl, extEl);
                    if (model != null)
                    {
                        worksheetModel.Shapes.Add(model);
                    }
                    continue;
                }

                var cxnSp = anchor.Element(XdrNs + "cxnSp");
                if (cxnSp != null)
                {
                    var model = ParseOneCellConnector(cxnSp, fromEl, extEl);
                    if (model != null)
                    {
                        worksheetModel.Shapes.Add(model);
                    }
                }
            }
        }

        private static ShapeModel ParseShape(XElement sp, XElement fromEl, XElement toEl)
        {
            var model = new ShapeModel
            {
                Name = GetShapeName(sp),
                UpperLeftColumn = ParseAnchorInt(fromEl.Element(XdrNs + "col")),
                UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff")),
                UpperLeftRow = ParseAnchorInt(fromEl.Element(XdrNs + "row")),
                UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff")),
                LowerRightColumn = ParseAnchorInt(toEl.Element(XdrNs + "col")),
                LowerRightColumnOffset = ParseAnchorLong(toEl.Element(XdrNs + "colOff")),
                LowerRightRow = ParseAnchorInt(toEl.Element(XdrNs + "row")),
                LowerRightRowOffset = ParseAnchorLong(toEl.Element(XdrNs + "rowOff")),
                GeometryType = GetGeometryType(sp),
                RawStyleXml = GetRawXml(sp.Element(XdrNs + "style")),
                RawTxBodyXml = GetRawXml(sp.Element(XdrNs + "txBody")),
                RawElementXml = sp.ToString(),
            };
            LoadShapeSpPrExtents(sp, model);
            return model;
        }

        private static ShapeModel ParseOneCellShape(XElement sp, XElement fromEl, XElement extEl)
        {
            var fromCol = ParseAnchorInt(fromEl.Element(XdrNs + "col"));
            var fromRow = ParseAnchorInt(fromEl.Element(XdrNs + "row"));
            long cx = 0;
            long cy = 0;

            if (extEl != null)
            {
                long.TryParse((string)extEl.Attribute("cx") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
                long.TryParse((string)extEl.Attribute("cy") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
            }

            var colSpan = cx > 0 ? (int)(cx / 609600L) + 1 : 1;
            var rowSpan = cy > 0 ? (int)(cy / 190500L) + 1 : 1;

            return new ShapeModel
            {
                Name = GetShapeName(sp),
                UpperLeftColumn = fromCol,
                UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff")),
                UpperLeftRow = fromRow,
                UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff")),
                LowerRightColumn = fromCol + colSpan,
                LowerRightRow = fromRow + rowSpan,
                ExtentCx = cx,
                ExtentCy = cy,
                GeometryType = GetGeometryType(sp),
                RawStyleXml = GetRawXml(sp.Element(XdrNs + "style")),
                RawTxBodyXml = GetRawXml(sp.Element(XdrNs + "txBody")),
                RawElementXml = sp.ToString(),
            };
        }

        private static ShapeModel ParseConnector(XElement cxnSp, XElement fromEl, XElement toEl)
        {
            var model = new ShapeModel
            {
                Name = GetConnectorName(cxnSp),
                UpperLeftColumn = ParseAnchorInt(fromEl.Element(XdrNs + "col")),
                UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff")),
                UpperLeftRow = ParseAnchorInt(fromEl.Element(XdrNs + "row")),
                UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff")),
                LowerRightColumn = ParseAnchorInt(toEl.Element(XdrNs + "col")),
                LowerRightColumnOffset = ParseAnchorLong(toEl.Element(XdrNs + "colOff")),
                LowerRightRow = ParseAnchorInt(toEl.Element(XdrNs + "row")),
                LowerRightRowOffset = ParseAnchorLong(toEl.Element(XdrNs + "rowOff")),
                GeometryType = GetGeometryType(cxnSp),
                RawElementXml = cxnSp.ToString(),
            };
            return model;
        }

        private static ShapeModel ParseOneCellConnector(XElement cxnSp, XElement fromEl, XElement extEl)
        {
            var fromCol = ParseAnchorInt(fromEl.Element(XdrNs + "col"));
            var fromRow = ParseAnchorInt(fromEl.Element(XdrNs + "row"));
            long cx = 0;
            long cy = 0;

            if (extEl != null)
            {
                long.TryParse((string)extEl.Attribute("cx") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
                long.TryParse((string)extEl.Attribute("cy") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
            }

            var colSpan = cx > 0 ? (int)(cx / 609600L) + 1 : 1;
            var rowSpan = cy > 0 ? (int)(cy / 190500L) + 1 : 1;

            return new ShapeModel
            {
                Name = GetConnectorName(cxnSp),
                UpperLeftColumn = fromCol,
                UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff")),
                UpperLeftRow = fromRow,
                UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff")),
                LowerRightColumn = fromCol + colSpan,
                LowerRightRow = fromRow + rowSpan,
                ExtentCx = cx,
                ExtentCy = cy,
                GeometryType = GetGeometryType(cxnSp),
                RawElementXml = cxnSp.ToString(),
            };
        }

        private static ShapeModel ParseGroupShape(XElement grpSp, XElement fromEl, XElement toEl)
        {
            return new ShapeModel
            {
                Name = GetGroupShapeName(grpSp),
                UpperLeftColumn = ParseAnchorInt(fromEl.Element(XdrNs + "col")),
                UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff")),
                UpperLeftRow = ParseAnchorInt(fromEl.Element(XdrNs + "row")),
                UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff")),
                LowerRightColumn = ParseAnchorInt(toEl.Element(XdrNs + "col")),
                LowerRightColumnOffset = ParseAnchorLong(toEl.Element(XdrNs + "colOff")),
                LowerRightRow = ParseAnchorInt(toEl.Element(XdrNs + "row")),
                LowerRightRowOffset = ParseAnchorLong(toEl.Element(XdrNs + "rowOff")),
                RawElementXml = grpSp.ToString(),
            };
        }

        private static ShapeModel ParseOneCellGroupShape(XElement grpSp, XElement fromEl, XElement extEl)
        {
            var fromCol = ParseAnchorInt(fromEl.Element(XdrNs + "col"));
            var fromRow = ParseAnchorInt(fromEl.Element(XdrNs + "row"));
            long cx = 0;
            long cy = 0;

            if (extEl != null)
            {
                long.TryParse((string)extEl.Attribute("cx") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
                long.TryParse((string)extEl.Attribute("cy") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
            }

            var colSpan = cx > 0 ? (int)(cx / 609600L) + 1 : 1;
            var rowSpan = cy > 0 ? (int)(cy / 190500L) + 1 : 1;

            return new ShapeModel
            {
                Name = GetGroupShapeName(grpSp),
                UpperLeftColumn = fromCol,
                UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff")),
                UpperLeftRow = fromRow,
                UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff")),
                LowerRightColumn = fromCol + colSpan,
                LowerRightRow = fromRow + rowSpan,
                ExtentCx = cx,
                ExtentCy = cy,
                RawElementXml = grpSp.ToString(),
            };
        }

        private static string GetGroupShapeName(XElement grpSp)
        {
            var nvGrpSpPr = grpSp.Element(XdrNs + "nvGrpSpPr");
            if (nvGrpSpPr == null)
            {
                return string.Empty;
            }

            var cNvPr = nvGrpSpPr.Element(XdrNs + "cNvPr");
            return (string)cNvPr?.Attribute("name") ?? string.Empty;
        }

        private static string GetConnectorName(XElement cxnSp)
        {
            var nvCxnSpPr = cxnSp.Element(XdrNs + "nvCxnSpPr");
            if (nvCxnSpPr == null)
            {
                return string.Empty;
            }

            var cNvPr = nvCxnSpPr.Element(XdrNs + "cNvPr");
            return (string)cNvPr?.Attribute("name") ?? string.Empty;
        }

        private static string GetShapeName(XElement sp)
        {
            var nvSpPr = sp.Element(XdrNs + "nvSpPr");
            if (nvSpPr == null)
            {
                return string.Empty;
            }

            var cNvPr = nvSpPr.Element(XdrNs + "cNvPr");
            return (string)cNvPr?.Attribute("name") ?? string.Empty;
        }

        private static string GetGeometryType(XElement sp)
        {
            var spPr = sp.Element(XdrNs + "spPr");
            if (spPr == null)
            {
                return "rect";
            }

            var prstGeom = spPr.Element(ANs + "prstGeom");
            return (string)prstGeom?.Attribute("prst") ?? "rect";
        }

        private static string GetRawXml(XElement element)
        {
            return element == null ? null : element.ToString();
        }

        private static void LoadShapeSpPrExtents(XElement sp, ShapeModel model)
        {
            var spPr = sp.Element(XdrNs + "spPr");
            if (spPr == null)
            {
                return;
            }

            var xfrm = spPr.Element(ANs + "xfrm");
            if (xfrm == null)
            {
                return;
            }

            var ext = xfrm.Element(ANs + "ext");
            if (ext == null)
            {
                return;
            }

            long cx;
            long cy;
            long.TryParse((string)ext.Attribute("cx") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
            long.TryParse((string)ext.Attribute("cy") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
            model.ExtentCx = cx;
            model.ExtentCy = cy;
        }

        private static PictureModel ParsePicture(XElement pic, XElement fromEl, XElement toEl, IReadOnlyDictionary<string, string> imageTargets, ZipArchive archive, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var rId = GetBlipRId(pic);
            if (string.IsNullOrEmpty(rId))
            {
                return null;
            }

            var name = GetPictureName(pic);
            var imageData = LoadImageData(rId, imageTargets, archive, diagnostics, options, sheetName, name);
            if (imageData == null)
            {
                return null;
            }

            var model = new PictureModel
            {
                Name = name,
                UpperLeftColumn = ParseAnchorInt(fromEl.Element(XdrNs + "col")),
                UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff")),
                UpperLeftRow = ParseAnchorInt(fromEl.Element(XdrNs + "row")),
                UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff")),
                LowerRightColumn = ParseAnchorInt(toEl.Element(XdrNs + "col")),
                LowerRightColumnOffset = ParseAnchorLong(toEl.Element(XdrNs + "colOff")),
                LowerRightRow = ParseAnchorInt(toEl.Element(XdrNs + "row")),
                LowerRightRowOffset = ParseAnchorLong(toEl.Element(XdrNs + "rowOff")),
                ImageData = imageData,
                ImageExtension = Picture.DetectExtension(imageData),
                OriginalRId = rId,
            };

            LoadSpPrExtents(pic, model);
            return model;
        }

        private static PictureModel ParseOneCellPicture(XElement pic, XElement fromEl, XElement extEl, IReadOnlyDictionary<string, string> imageTargets, ZipArchive archive, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var rId = GetBlipRId(pic);
            if (string.IsNullOrEmpty(rId))
            {
                return null;
            }

            var name = GetPictureName(pic);
            var imageData = LoadImageData(rId, imageTargets, archive, diagnostics, options, sheetName, name);
            if (imageData == null)
            {
                return null;
            }

            var fromCol = ParseAnchorInt(fromEl.Element(XdrNs + "col"));
            var fromRow = ParseAnchorInt(fromEl.Element(XdrNs + "row"));
            long cx = 0;
            long cy = 0;

            if (extEl != null)
            {
                long.TryParse((string)extEl.Attribute("cx") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
                long.TryParse((string)extEl.Attribute("cy") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
            }

            var colSpan = cx > 0 ? (int)(cx / 609600L) + 1 : 1;
            var rowSpan = cy > 0 ? (int)(cy / 190500L) + 1 : 1;

            var model = new PictureModel
            {
                Name = name,
                UpperLeftColumn = fromCol,
                UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff")),
                UpperLeftRow = fromRow,
                UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff")),
                LowerRightColumn = fromCol + colSpan,
                LowerRightRow = fromRow + rowSpan,
                ExtentCx = cx,
                ExtentCy = cy,
                ImageData = imageData,
                ImageExtension = Picture.DetectExtension(imageData),
                OriginalRId = rId,
            };
            return model;
        }

        private static string GetBlipRId(XElement pic)
        {
            var blipFill = pic.Element(XdrNs + "blipFill");
            if (blipFill == null)
            {
                return null;
            }

            var blip = blipFill.Element(ANs + "blip");
            if (blip == null)
            {
                return null;
            }

            return (string)blip.Attribute(RelationshipNs + "embed");
        }

        private static string GetPictureName(XElement pic)
        {
            var nvPicPr = pic.Element(XdrNs + "nvPicPr");
            if (nvPicPr == null)
            {
                return string.Empty;
            }

            var cNvPr = nvPicPr.Element(XdrNs + "cNvPr");
            return (string)cNvPr?.Attribute("name") ?? string.Empty;
        }

        private static byte[] LoadImageData(string rId, IReadOnlyDictionary<string, string> imageTargets, ZipArchive archive, LoadDiagnostics diagnostics, LoadOptions options, string sheetName, string pictureName)
        {
            string mediaUri;
            if (!imageTargets.TryGetValue(rId, out mediaUri))
            {
                AddIssue(diagnostics, options, new LoadIssue("PIC-L001", DiagnosticSeverity.LossyRecoverable, "Image relationship '" + rId + "' for picture '" + pictureName + "' could not be resolved; picture was dropped.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                });
                return null;
            }

            var mediaEntry = GetEntry(archive, mediaUri);
            if (mediaEntry == null)
            {
                AddIssue(diagnostics, options, new LoadIssue("PIC-L001", DiagnosticSeverity.LossyRecoverable, "Image media file '" + mediaUri + "' was not found in the package; picture was dropped.", repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                });
                return null;
            }

            using (var stream = mediaEntry.Open())
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                return ms.ToArray();
            }
        }

        private static void LoadSpPrExtents(XElement pic, PictureModel model)
        {
            var spPr = pic.Element(XdrNs + "spPr");
            if (spPr == null)
            {
                return;
            }

            var xfrm = spPr.Element(ANs + "xfrm");
            if (xfrm == null)
            {
                return;
            }

            var ext = xfrm.Element(ANs + "ext");
            if (ext == null)
            {
                return;
            }

            long cx, cy;
            long.TryParse((string)ext.Attribute("cx") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
            long.TryParse((string)ext.Attribute("cy") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
            model.ExtentCx = cx;
            model.ExtentCy = cy;
        }

        private static IReadOnlyDictionary<string, string> LoadDrawingImageTargets(ZipArchive archive, string drawingUri)
        {
            return LoadDrawingTargetsByType(archive, drawingUri, ImageRelationshipType);
        }

        private sealed class ChartTarget
        {
            public string Uri;
            public bool IsChartEx;
        }

        private static IReadOnlyDictionary<string, ChartTarget> LoadDrawingChartTargets(ZipArchive archive, string drawingUri)
        {
            var relsUri = GetDrawingRelsUri(drawingUri);
            var entry = GetEntry(archive, relsUri);
            if (entry == null)
            {
                return new Dictionary<string, ChartTarget>(StringComparer.OrdinalIgnoreCase);
            }

            var document = LoadDocument(entry);
            var targets = new Dictionary<string, ChartTarget>(StringComparer.OrdinalIgnoreCase);
            foreach (var rel in document.Root != null ? document.Root.Elements(PackageRelationshipNs + "Relationship") : new XElement[0])
            {
                var id = (string)rel.Attribute("Id");
                var type = (string)rel.Attribute("Type");
                var target = (string)rel.Attribute("Target");
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target))
                {
                    continue;
                }

                var isStandard = string.Equals(type, ChartRelationshipType, StringComparison.OrdinalIgnoreCase);
                var isChartEx = string.Equals(type, ChartExRelationshipType, StringComparison.OrdinalIgnoreCase);
                if (!isStandard && !isChartEx)
                {
                    continue;
                }

                targets[id] = new ChartTarget { Uri = ResolvePartUri(drawingUri, target), IsChartEx = isChartEx };
            }

            return targets;
        }

        private static IReadOnlyDictionary<string, string> LoadDrawingTargetsByType(ZipArchive archive, string drawingUri, string relationshipType)
        {
            var relsUri = GetDrawingRelsUri(drawingUri);
            var entry = GetEntry(archive, relsUri);
            if (entry == null)
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            var document = LoadDocument(entry);
            var targets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var rel in document.Root != null ? document.Root.Elements(PackageRelationshipNs + "Relationship") : new XElement[0])
            {
                var id = (string)rel.Attribute("Id");
                var type = (string)rel.Attribute("Type");
                var target = (string)rel.Attribute("Target");
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target))
                {
                    continue;
                }

                if (!string.Equals(type, relationshipType, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                targets[id] = ResolvePartUri(drawingUri, target);
            }

            return targets;
        }

        private static string FindDrawingUri(ZipArchive archive, string worksheetUri)
        {
            var relsUri = GetWorksheetRelsUri(worksheetUri);
            var entry = GetEntry(archive, relsUri);
            if (entry == null)
            {
                return string.Empty;
            }

            var document = LoadDocument(entry);
            if (document.Root == null)
            {
                return string.Empty;
            }

            foreach (var rel in document.Root.Elements(PackageRelationshipNs + "Relationship"))
            {
                var type = (string)rel.Attribute("Type");
                var target = (string)rel.Attribute("Target");
                if (string.Equals(type, DrawingRelationshipType, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(target))
                {
                    return ResolvePartUri(worksheetUri, target);
                }
            }

            return string.Empty;
        }

        private static int ParseAnchorInt(XElement element)
        {
            if (element == null)
            {
                return 0;
            }

            int value;
            int.TryParse(element.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
            return value;
        }

        private static long ParseAnchorLong(XElement element)
        {
            if (element == null)
            {
                return 0L;
            }

            long value;
            long.TryParse(element.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
            return value;
        }

        private static string GetWorksheetRelsUri(string worksheetUri)
        {
            var normalized = worksheetUri.TrimStart('/');
            var slashIndex = normalized.LastIndexOf('/');
            var directory = slashIndex >= 0 ? normalized.Substring(0, slashIndex + 1) : string.Empty;
            var fileName = slashIndex >= 0 ? normalized.Substring(slashIndex + 1) : normalized;
            return "/" + directory + "_rels/" + fileName + ".rels";
        }

        private static string GetDrawingRelsUri(string drawingUri)
        {
            var normalized = drawingUri.TrimStart('/');
            var slashIndex = normalized.LastIndexOf('/');
            var directory = slashIndex >= 0 ? normalized.Substring(0, slashIndex + 1) : string.Empty;
            var fileName = slashIndex >= 0 ? normalized.Substring(slashIndex + 1) : normalized;
            return "/" + directory + "_rels/" + fileName + ".rels";
        }

        // --- Chart loading ---

        private static readonly XNamespace McNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        private const string ChartExGraphicDataUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";

        private static void LoadTwoCellAnchorCharts(WorksheetModel worksheetModel, XElement drawingRoot,
            IReadOnlyDictionary<string, ChartTarget> chartTargets, ZipArchive archive,
            LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            foreach (var anchor in drawingRoot.Elements(XdrNs + "twoCellAnchor"))
            {
                var fromEl = anchor.Element(XdrNs + "from");
                var toEl = anchor.Element(XdrNs + "to");
                if (fromEl == null || toEl == null)
                {
                    continue;
                }

                // Standard chart: direct graphicFrame child
                var gf = anchor.Element(XdrNs + "graphicFrame");
                if (gf != null && IsChartGraphicFrame(gf))
                {
                    var model = ParseChart(gf, null, fromEl, toEl, null, chartTargets, archive, diagnostics, options, sheetName);
                    if (model != null) worksheetModel.Charts.Add(model);
                    continue;
                }

                // ChartEx: mc:AlternateContent wrapping a graphicFrame with chartex URI
                var altContent = anchor.Element(McNamespace + "AlternateContent");
                if (altContent != null)
                {
                    var choice = altContent.Element(McNamespace + "Choice");
                    var choiceGf = choice?.Element(XdrNs + "graphicFrame");
                    if (choiceGf != null && IsChartExGraphicFrame(choiceGf))
                    {
                        var model = ParseChart(choiceGf, altContent, fromEl, toEl, null, chartTargets, archive, diagnostics, options, sheetName);
                        if (model != null) worksheetModel.Charts.Add(model);
                    }
                }
            }
        }

        private static void LoadOneCellAnchorCharts(WorksheetModel worksheetModel, XElement drawingRoot,
            IReadOnlyDictionary<string, ChartTarget> chartTargets, ZipArchive archive,
            LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            foreach (var anchor in drawingRoot.Elements(XdrNs + "oneCellAnchor"))
            {
                var fromEl = anchor.Element(XdrNs + "from");
                var extEl = anchor.Element(XdrNs + "ext");
                if (fromEl == null)
                {
                    continue;
                }

                var gf = anchor.Element(XdrNs + "graphicFrame");
                if (gf != null && IsChartGraphicFrame(gf))
                {
                    var model = ParseChart(gf, null, fromEl, null, extEl, chartTargets, archive, diagnostics, options, sheetName);
                    if (model != null) worksheetModel.Charts.Add(model);
                    continue;
                }

                var altContent = anchor.Element(McNamespace + "AlternateContent");
                if (altContent != null)
                {
                    var choice = altContent.Element(McNamespace + "Choice");
                    var choiceGf = choice?.Element(XdrNs + "graphicFrame");
                    if (choiceGf != null && IsChartExGraphicFrame(choiceGf))
                    {
                        var model = ParseChart(choiceGf, altContent, fromEl, null, extEl, chartTargets, archive, diagnostics, options, sheetName);
                        if (model != null) worksheetModel.Charts.Add(model);
                    }
                }
            }
        }

        private static bool IsChartGraphicFrame(XElement graphicFrame)
        {
            var graphicData = graphicFrame.Element(ANs + "graphic")?.Element(ANs + "graphicData");
            if (graphicData == null)
            {
                return false;
            }

            var uri = (string)graphicData.Attribute("uri");
            return string.Equals(uri, ChartGraphicDataUri, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsChartExGraphicFrame(XElement graphicFrame)
        {
            var graphicData = graphicFrame.Element(ANs + "graphic")?.Element(ANs + "graphicData");
            if (graphicData == null)
            {
                return false;
            }

            var uri = (string)graphicData.Attribute("uri");
            return string.Equals(uri, ChartExGraphicDataUri, StringComparison.OrdinalIgnoreCase);
        }

        private static readonly XNamespace ChartExNs = "http://schemas.microsoft.com/office/drawing/2014/chartex";

        private static ChartModel ParseChart(XElement graphicFrame, XElement rawContainerElement,
            XElement fromEl, XElement toEl, XElement extEl,
            IReadOnlyDictionary<string, ChartTarget> chartTargets, ZipArchive archive,
            LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var graphicData = graphicFrame.Element(ANs + "graphic")?.Element(ANs + "graphicData");

            // Standard chart: c:chart element; ChartEx: cx:chart element
            var chartEl = graphicData?.Element(ChartNs + "chart")
                       ?? graphicData?.Element(ChartExNs + "chart");
            var rId = chartEl != null ? (string)chartEl.Attribute(RelationshipNs + "id") : null;
            if (string.IsNullOrEmpty(rId))
            {
                return null;
            }

            ChartTarget chartTarget;
            if (!chartTargets.TryGetValue(rId, out chartTarget))
            {
                AddIssue(diagnostics, options, new LoadIssue("CHT-L001", DiagnosticSeverity.LossyRecoverable,
                    "Chart relationship '" + rId + "' could not be resolved; chart was dropped.",
                    repairApplied: true, dataLossRisk: true)
                {
                    SheetName = sheetName,
                });
                return null;
            }

            var chartUri = chartTarget.Uri;
            var isChartEx = chartTarget.IsChartEx;

            var chartEntry = GetEntry(archive, chartUri);
            if (chartEntry == null)
            {
                AddIssue(diagnostics, options, new LoadIssue("CHT-R001", DiagnosticSeverity.Recoverable,
                    (isChartEx ? "ChartEx" : "Chart") + " part '" + chartUri + "' was not found; chart was skipped.",
                    repairApplied: true)
                {
                    SheetName = sheetName,
                });
                return null;
            }

            string rawChartXml;
            using (var stream = chartEntry.Open())
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                rawChartXml = reader.ReadToEnd();
            }

            var model = new ChartModel
            {
                Name = GetChartName(graphicFrame),
                RawChartXml = rawChartXml,
                ChartType = DetectChartType(rawChartXml),
                IsChartEx = isChartEx,
                OriginalRId = rId,
                RawGraphicFrameXml = rawContainerElement != null ? rawContainerElement.ToString() : null,
            };

            if (toEl != null)
            {
                model.UpperLeftColumn = ParseAnchorInt(fromEl.Element(XdrNs + "col"));
                model.UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff"));
                model.UpperLeftRow = ParseAnchorInt(fromEl.Element(XdrNs + "row"));
                model.UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff"));
                model.LowerRightColumn = ParseAnchorInt(toEl.Element(XdrNs + "col"));
                model.LowerRightColumnOffset = ParseAnchorLong(toEl.Element(XdrNs + "colOff"));
                model.LowerRightRow = ParseAnchorInt(toEl.Element(XdrNs + "row"));
                model.LowerRightRowOffset = ParseAnchorLong(toEl.Element(XdrNs + "rowOff"));
            }
            else if (extEl != null)
            {
                var fromCol = ParseAnchorInt(fromEl.Element(XdrNs + "col"));
                var fromRow = ParseAnchorInt(fromEl.Element(XdrNs + "row"));
                long cx = 0;
                long cy = 0;
                long.TryParse((string)extEl.Attribute("cx") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
                long.TryParse((string)extEl.Attribute("cy") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
                model.UpperLeftColumn = fromCol;
                model.UpperLeftColumnOffset = ParseAnchorLong(fromEl.Element(XdrNs + "colOff"));
                model.UpperLeftRow = fromRow;
                model.UpperLeftRowOffset = ParseAnchorLong(fromEl.Element(XdrNs + "rowOff"));
                model.LowerRightColumn = fromCol + (cx > 0 ? (int)(cx / 609600L) + 1 : 1);
                model.LowerRightRow = fromRow + (cy > 0 ? (int)(cy / 190500L) + 1 : 1);
                model.ExtentCx = cx;
                model.ExtentCy = cy;
            }

            LoadChartCompanionFiles(model, chartUri, archive, diagnostics, options, sheetName);
            return model;
        }

        private static string GetChartName(XElement graphicFrame)
        {
            var nvGraphicFramePr = graphicFrame.Element(XdrNs + "nvGraphicFramePr");
            var cNvPr = nvGraphicFramePr?.Element(XdrNs + "cNvPr");
            return (string)cNvPr?.Attribute("name") ?? string.Empty;
        }

        private static ChartType DetectChartType(string rawXml)
        {
            if (string.IsNullOrEmpty(rawXml))
            {
                return ChartType.Unknown;
            }

            if (rawXml.IndexOf("<c:barChart", StringComparison.Ordinal) >= 0)
            {
                return rawXml.IndexOf("barDir val=\"bar\"", StringComparison.Ordinal) >= 0
                    ? ChartType.Bar
                    : ChartType.Column;
            }

            if (rawXml.IndexOf("<c:bar3DChart", StringComparison.Ordinal) >= 0)
            {
                return rawXml.IndexOf("barDir val=\"bar\"", StringComparison.Ordinal) >= 0
                    ? ChartType.Bar3D
                    : ChartType.Column3D;
            }

            if (rawXml.IndexOf("<c:lineChart", StringComparison.Ordinal) >= 0) return ChartType.Line;
            if (rawXml.IndexOf("<c:line3DChart", StringComparison.Ordinal) >= 0) return ChartType.Line3D;
            if (rawXml.IndexOf("<c:areaChart", StringComparison.Ordinal) >= 0) return ChartType.Area;
            if (rawXml.IndexOf("<c:area3DChart", StringComparison.Ordinal) >= 0) return ChartType.Area3D;
            if (rawXml.IndexOf("<c:pie3DChart", StringComparison.Ordinal) >= 0) return ChartType.Pie3D;
            if (rawXml.IndexOf("<c:pieChart", StringComparison.Ordinal) >= 0) return ChartType.Pie;
            if (rawXml.IndexOf("<c:doughnutChart", StringComparison.Ordinal) >= 0) return ChartType.Doughnut;
            if (rawXml.IndexOf("<c:scatterChart", StringComparison.Ordinal) >= 0) return ChartType.Scatter;
            if (rawXml.IndexOf("<c:bubbleChart", StringComparison.Ordinal) >= 0) return ChartType.Bubble;
            if (rawXml.IndexOf("<c:radarChart", StringComparison.Ordinal) >= 0) return ChartType.Radar;
            if (rawXml.IndexOf("<c:stockChart", StringComparison.Ordinal) >= 0) return ChartType.Stock;
            if (rawXml.IndexOf("<c:surface3DChart", StringComparison.Ordinal) >= 0)
            {
                return rawXml.IndexOf("wireframe val=\"1\"", StringComparison.Ordinal) >= 0
                    ? ChartType.SurfaceWireframe3D
                    : ChartType.Surface3D;
            }

            if (rawXml.IndexOf("<c:surfaceChart", StringComparison.Ordinal) >= 0) return ChartType.Contour;

            return ChartType.Unknown;
        }

        private static void LoadChartCompanionFiles(ChartModel model, string chartUri, ZipArchive archive,
            LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var normalized = chartUri.TrimStart('/');
            var slashIndex = normalized.LastIndexOf('/');
            var directory = slashIndex >= 0 ? normalized.Substring(0, slashIndex + 1) : string.Empty;
            var fileName = slashIndex >= 0 ? normalized.Substring(slashIndex + 1) : normalized;
            var chartRelsUri = "/" + directory + "_rels/" + fileName + ".rels";

            var relsEntry = GetEntry(archive, chartRelsUri);
            if (relsEntry == null)
            {
                return;
            }

            var relsDocument = LoadDocument(relsEntry);
            if (relsDocument.Root == null)
            {
                return;
            }

            foreach (var rel in relsDocument.Root.Elements(PackageRelationshipNs + "Relationship"))
            {
                var relId = (string)rel.Attribute("Id");
                var type = (string)rel.Attribute("Type");
                var target = (string)rel.Attribute("Target");
                if (string.IsNullOrEmpty(relId) || string.IsNullOrEmpty(target))
                {
                    continue;
                }

                var isXmlCompanion =
                    string.Equals(type, ChartStyleRelationshipType, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(type, ChartColorStyleRelationshipType, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(type, ChartUserShapesRelationshipType, StringComparison.OrdinalIgnoreCase);

                if (isXmlCompanion)
                {
                    var companionUri = ResolvePartUri(chartUri, target);
                    var companionEntry = GetEntry(archive, companionUri);
                    if (companionEntry == null)
                    {
                        continue;
                    }

                    string rawContent;
                    using (var stream = companionEntry.Open())
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        rawContent = reader.ReadToEnd();
                    }

                    // Store the original relative target path so cross-directory paths (e.g.
                    // "../drawings/drawing4.xml") can be resolved correctly on save without renumbering.
                    model.CompanionFiles.Add(new ChartCompanionFile
                    {
                        RelationshipId = relId,
                        RelationshipType = type,
                        FileName = target,
                        RawContent = rawContent,
                    });
                }
                else if (string.Equals(type, ImageRelationshipType, StringComparison.OrdinalIgnoreCase))
                {
                    AddIssue(diagnostics, options, new LoadIssue("CHT-R002", DiagnosticSeverity.Recoverable,
                        "Chart '" + model.Name + "' contains an embedded image companion; this is not yet supported and the image was not loaded.",
                        repairApplied: false)
                    {
                        SheetName = sheetName,
                    });
                }
            }
        }
    }
}
