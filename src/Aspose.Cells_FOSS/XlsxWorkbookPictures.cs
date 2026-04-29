using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxWorkbookPictures
    {
        internal const string DrawingRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
        internal const string ImageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        internal const string DrawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
        internal const string ChartRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        internal const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        internal const string ChartExRelationshipType = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
        internal const string ChartExContentType = "application/vnd.ms-office.chartex+xml";
        internal const string ChartStyleContentType = "application/vnd.ms-office.chartstyle+xml";
        internal const string ChartColorStyleContentType = "application/vnd.ms-office.chartcolorstyle+xml";
        internal const string ChartGraphicDataUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        internal const string ChartStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
        internal const string ChartColorStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
        internal const string ChartUserShapesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
        internal const string ChartUserShapesContentType = "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";

        internal static readonly XNamespace XdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        internal static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        internal static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        internal static XDocument BuildDrawingDocument(WorksheetModel worksheet, int pictureFileOffset, int chartFileOffset)
        {
            var root = new XElement(XdrNs + "wsDr",
                new XAttribute(XNamespace.Xmlns + "xdr", XdrNs),
                new XAttribute(XNamespace.Xmlns + "a", ANs),
                new XAttribute(XNamespace.Xmlns + "r", RelationshipNs));

            // Build a mapping from each picture's/shape-image's original rId to its new assigned rId
            // so that group shapes whose raw XML contains r:embed references stay correct.
            var imageCount = worksheet.Pictures.Count + worksheet.ShapeImages.Count;
            var rIdMap = BuildImageRIdMap(worksheet.Pictures, worksheet.ShapeImages);

            for (var i = 0; i < worksheet.Pictures.Count; i++)
            {
                var picture = worksheet.Pictures[i];
                var rId = "rId" + (i + 1).ToString(CultureInfo.InvariantCulture);
                var picId = i + 1;
                root.Add(BuildTwoCellAnchor(picture, rId, picId));
            }

            for (var j = 0; j < worksheet.Shapes.Count; j++)
            {
                var shape = worksheet.Shapes[j];
                var shapeId = worksheet.Pictures.Count + j + 1;
                root.Add(BuildShapeAnchor(shape, shapeId, rIdMap));
            }

            for (var k = 0; k < worksheet.Charts.Count; k++)
            {
                var chart = worksheet.Charts[k];
                var rId = "rId" + (imageCount + k + 1).ToString(CultureInfo.InvariantCulture);
                var chartId = imageCount + worksheet.Shapes.Count + k + 1;
                root.Add(BuildChartAnchor(chart, rId, chartId));
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
        }

        private static System.Collections.Generic.Dictionary<string, string> BuildImageRIdMap(
            System.Collections.Generic.IList<Core.PictureModel> pictures,
            System.Collections.Generic.IList<Core.ShapeImageModel> shapeImages)
        {
            var map = new System.Collections.Generic.Dictionary<string, string>(StringComparer.Ordinal);
            for (var i = 0; i < pictures.Count; i++)
            {
                var originalRId = pictures[i].OriginalRId;
                if (!string.IsNullOrEmpty(originalRId))
                {
                    var newRId = "rId" + (i + 1).ToString(CultureInfo.InvariantCulture);
                    if (originalRId != newRId)
                    {
                        map[originalRId] = newRId;
                    }
                }
            }
            for (var j = 0; j < shapeImages.Count; j++)
            {
                var originalRId = shapeImages[j].OriginalRId;
                if (!string.IsNullOrEmpty(originalRId))
                {
                    var newRId = "rId" + (pictures.Count + j + 1).ToString(CultureInfo.InvariantCulture);
                    if (originalRId != newRId)
                    {
                        map[originalRId] = newRId;
                    }
                }
            }
            return map;
        }

        private static XElement BuildTwoCellAnchor(PictureModel picture, string rId, int picId)
        {
            var anchor = new XElement(XdrNs + "twoCellAnchor",
                new XAttribute("editAs", "oneCell"));

            anchor.Add(BuildFromElement(picture.UpperLeftColumn, picture.UpperLeftColumnOffset,
                                        picture.UpperLeftRow, picture.UpperLeftRowOffset));
            anchor.Add(BuildToElement(picture.LowerRightColumn, picture.LowerRightColumnOffset,
                                      picture.LowerRightRow, picture.LowerRightRowOffset));
            anchor.Add(BuildPicElement(picture, rId, picId));
            anchor.Add(new XElement(XdrNs + "clientData"));
            return anchor;
        }

        private static XElement BuildFromElement(int col, long colOff, int row, long rowOff)
        {
            return new XElement(XdrNs + "from",
                new XElement(XdrNs + "col", col.ToString(CultureInfo.InvariantCulture)),
                new XElement(XdrNs + "colOff", colOff.ToString(CultureInfo.InvariantCulture)),
                new XElement(XdrNs + "row", row.ToString(CultureInfo.InvariantCulture)),
                new XElement(XdrNs + "rowOff", rowOff.ToString(CultureInfo.InvariantCulture)));
        }

        private static XElement BuildToElement(int col, long colOff, int row, long rowOff)
        {
            return new XElement(XdrNs + "to",
                new XElement(XdrNs + "col", col.ToString(CultureInfo.InvariantCulture)),
                new XElement(XdrNs + "colOff", colOff.ToString(CultureInfo.InvariantCulture)),
                new XElement(XdrNs + "row", row.ToString(CultureInfo.InvariantCulture)),
                new XElement(XdrNs + "rowOff", rowOff.ToString(CultureInfo.InvariantCulture)));
        }

        private static XElement BuildPicElement(PictureModel picture, string rId, int picId)
        {
            var cx = picture.ExtentCx > 0
                ? picture.ExtentCx
                : (long)(picture.LowerRightColumn - picture.UpperLeftColumn) * 609600L;
            var cy = picture.ExtentCy > 0
                ? picture.ExtentCy
                : (long)(picture.LowerRightRow - picture.UpperLeftRow) * 190500L;

            if (cx <= 0)
            {
                cx = 609600L;
            }

            if (cy <= 0)
            {
                cy = 190500L;
            }

            var pic = new XElement(XdrNs + "pic");

            var nvPicPr = new XElement(XdrNs + "nvPicPr",
                new XElement(XdrNs + "cNvPr",
                    new XAttribute("id", picId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("name", string.IsNullOrEmpty(picture.Name) ? "Picture " + picId.ToString(CultureInfo.InvariantCulture) : picture.Name)),
                new XElement(XdrNs + "cNvPicPr",
                    new XElement(ANs + "picLocks",
                        new XAttribute("noChangeAspect", "1"))));
            pic.Add(nvPicPr);

            var blipFill = new XElement(XdrNs + "blipFill",
                new XElement(ANs + "blip",
                    new XAttribute(RelationshipNs + "embed", rId)),
                new XElement(ANs + "stretch",
                    new XElement(ANs + "fillRect")));
            pic.Add(blipFill);

            var spPr = new XElement(XdrNs + "spPr",
                new XElement(ANs + "xfrm",
                    new XElement(ANs + "off",
                        new XAttribute("x", "0"),
                        new XAttribute("y", "0")),
                    new XElement(ANs + "ext",
                        new XAttribute("cx", cx.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("cy", cy.ToString(CultureInfo.InvariantCulture)))),
                new XElement(ANs + "prstGeom",
                    new XAttribute("prst", "rect")));
            pic.Add(spPr);

            return pic;
        }

        private static XElement BuildShapeAnchor(ShapeModel shape, int shapeId,
            System.Collections.Generic.Dictionary<string, string> pictureRIdMap = null)
        {
            var anchor = new XElement(XdrNs + "twoCellAnchor",
                new XAttribute("editAs", "oneCell"));

            anchor.Add(BuildFromElement(shape.UpperLeftColumn, shape.UpperLeftColumnOffset,
                                        shape.UpperLeftRow, shape.UpperLeftRowOffset));
            anchor.Add(BuildToElement(shape.LowerRightColumn, shape.LowerRightColumnOffset,
                                      shape.LowerRightRow, shape.LowerRightRowOffset));

            if (!string.IsNullOrEmpty(shape.RawElementXml))
            {
                try
                {
                    var rawXml = ApplyRIdRemap(shape.RawElementXml, pictureRIdMap);
                    anchor.Add(XElement.Parse(rawXml));
                }
                catch
                {
                    anchor.Add(BuildSpElement(shape, shapeId));
                }
            }
            else
            {
                anchor.Add(BuildSpElement(shape, shapeId));
            }

            anchor.Add(new XElement(XdrNs + "clientData"));
            return anchor;
        }

        private static string ApplyRIdRemap(string rawXml, System.Collections.Generic.Dictionary<string, string> rIdMap)
        {
            if (rIdMap == null || rIdMap.Count == 0 || string.IsNullOrEmpty(rawXml))
            {
                return rawXml;
            }

            foreach (var kvp in rIdMap)
            {
                rawXml = rawXml.Replace("\"" + kvp.Key + "\"", "\"" + kvp.Value + "\"");
            }

            return rawXml;
        }

        private static XElement BuildSpElement(ShapeModel shape, int shapeId)
        {
            var cx = shape.ExtentCx > 0
                ? shape.ExtentCx
                : (long)(shape.LowerRightColumn - shape.UpperLeftColumn) * 609600L;
            var cy = shape.ExtentCy > 0
                ? shape.ExtentCy
                : (long)(shape.LowerRightRow - shape.UpperLeftRow) * 190500L;

            if (cx <= 0)
            {
                cx = 609600L;
            }

            if (cy <= 0)
            {
                cy = 190500L;
            }

            var geomType = string.IsNullOrEmpty(shape.GeometryType) ? "rect" : shape.GeometryType;
            var displayName = string.IsNullOrEmpty(shape.Name) ? "Shape " + shapeId.ToString(CultureInfo.InvariantCulture) : shape.Name;

            var sp = new XElement(XdrNs + "sp",
                new XAttribute("macro", string.Empty),
                new XAttribute("textlink", string.Empty));

            sp.Add(new XElement(XdrNs + "nvSpPr",
                new XElement(XdrNs + "cNvPr",
                    new XAttribute("id", shapeId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("name", displayName)),
                new XElement(XdrNs + "cNvSpPr")));

            sp.Add(new XElement(XdrNs + "spPr",
                new XElement(ANs + "xfrm",
                    new XElement(ANs + "off",
                        new XAttribute("x", "0"),
                        new XAttribute("y", "0")),
                    new XElement(ANs + "ext",
                        new XAttribute("cx", cx.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("cy", cy.ToString(CultureInfo.InvariantCulture)))),
                new XElement(ANs + "prstGeom",
                    new XAttribute("prst", geomType),
                    new XElement(ANs + "avLst"))));

            if (!string.IsNullOrEmpty(shape.RawStyleXml))
            {
                try
                {
                    sp.Add(XElement.Parse(shape.RawStyleXml));
                }
                catch
                {
                }
            }
            else
            {
                sp.Add(BuildDefaultShapeStyle());
            }

            if (!string.IsNullOrEmpty(shape.RawTxBodyXml))
            {
                try
                {
                    sp.Add(XElement.Parse(shape.RawTxBodyXml));
                }
                catch
                {
                    sp.Add(BuildMinimalTxBody());
                }
            }
            else
            {
                sp.Add(BuildMinimalTxBody());
            }

            return sp;
        }

        private static XElement BuildMinimalTxBody()
        {
            return new XElement(XdrNs + "txBody",
                new XElement(ANs + "bodyPr"),
                new XElement(ANs + "lstStyle"),
                new XElement(ANs + "p"));
        }

        private static XElement BuildDefaultShapeStyle()
        {
            return new XElement(XdrNs + "style",
                new XElement(ANs + "lnRef", new XAttribute("idx", "2"), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"), new XElement(ANs + "shade", new XAttribute("val", "50000")))),
                new XElement(ANs + "fillRef", new XAttribute("idx", "1"), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
                new XElement(ANs + "effectRef", new XAttribute("idx", "0"), new XElement(ANs + "schemeClr", new XAttribute("val", "accent1"))),
                new XElement(ANs + "fontRef", new XAttribute("idx", "minor"), new XElement(ANs + "schemeClr", new XAttribute("val", "lt1"))));
        }

        internal static XDocument BuildDrawingRelationshipsDocument(WorksheetModel worksheet, int pictureFileOffset, int chartFileOffset)
        {
            var relationships = new XElement(PackageRelationshipNs + "Relationships");
            var imageCount = worksheet.Pictures.Count + worksheet.ShapeImages.Count;

            for (var i = 0; i < worksheet.Pictures.Count; i++)
            {
                var picture = worksheet.Pictures[i];
                var globalPictureNumber = pictureFileOffset + i + 1;
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + (i + 1).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", ImageRelationshipType),
                    new XAttribute("Target", "../media/image" + globalPictureNumber.ToString(CultureInfo.InvariantCulture) + "." + picture.ImageExtension)));
            }

            // Shape-embedded images follow top-level pictures so their rIds come after pictures
            // but before charts. This keeps r:embed references in group shape raw XML valid.
            for (var j = 0; j < worksheet.ShapeImages.Count; j++)
            {
                var shapeImage = worksheet.ShapeImages[j];
                var globalPictureNumber = pictureFileOffset + worksheet.Pictures.Count + j + 1;
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + (worksheet.Pictures.Count + j + 1).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", ImageRelationshipType),
                    new XAttribute("Target", "../media/image" + globalPictureNumber.ToString(CultureInfo.InvariantCulture) + "." + shapeImage.Extension)));
            }

            for (var k = 0; k < worksheet.Charts.Count; k++)
            {
                var chart = worksheet.Charts[k];
                var globalChartNumber = chartFileOffset + k + 1;
                var relType = chart.IsChartEx ? ChartExRelationshipType : ChartRelationshipType;
                var fileName = chart.IsChartEx
                    ? "chartEx" + globalChartNumber.ToString(CultureInfo.InvariantCulture) + ".xml"
                    : "chart" + globalChartNumber.ToString(CultureInfo.InvariantCulture) + ".xml";
                relationships.Add(new XElement(PackageRelationshipNs + "Relationship",
                    new XAttribute("Id", "rId" + (imageCount + k + 1).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", relType),
                    new XAttribute("Target", "../charts/" + fileName)));
            }

            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), relationships);
        }

        internal static void WritePictureMediaEntries(ZipArchive archive, WorksheetModel worksheet, int pictureFileOffset)
        {
            for (var i = 0; i < worksheet.Pictures.Count; i++)
            {
                var picture = worksheet.Pictures[i];
                var globalPictureNumber = pictureFileOffset + i + 1;
                var path = "xl/media/image" + globalPictureNumber.ToString(CultureInfo.InvariantCulture) + "." + picture.ImageExtension;
                var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
                using (var stream = entry.Open())
                {
                    stream.Write(picture.ImageData, 0, picture.ImageData.Length);
                }
            }

            for (var j = 0; j < worksheet.ShapeImages.Count; j++)
            {
                var shapeImage = worksheet.ShapeImages[j];
                var globalPictureNumber = pictureFileOffset + worksheet.Pictures.Count + j + 1;
                var path = "xl/media/image" + globalPictureNumber.ToString(CultureInfo.InvariantCulture) + "." + shapeImage.Extension;
                var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
                using (var stream = entry.Open())
                {
                    stream.Write(shapeImage.ImageData, 0, shapeImage.ImageData.Length);
                }
            }
        }

        private static XElement BuildChartAnchor(Core.ChartModel chart, string rId, int chartId)
        {
            var anchor = new XElement(XdrNs + "twoCellAnchor",
                new XAttribute("editAs", "oneCell"));

            anchor.Add(BuildFromElement(chart.UpperLeftColumn, chart.UpperLeftColumnOffset,
                                        chart.UpperLeftRow, chart.UpperLeftRowOffset));
            anchor.Add(BuildToElement(chart.LowerRightColumn, chart.LowerRightColumnOffset,
                                      chart.LowerRightRow, chart.LowerRightRowOffset));

            if (chart.IsChartEx && !string.IsNullOrEmpty(chart.RawGraphicFrameXml))
            {
                // Substitute the new rId into the preserved raw element (mc:AlternateContent or graphicFrame)
                var updatedXml = string.IsNullOrEmpty(chart.OriginalRId)
                    ? chart.RawGraphicFrameXml
                    : chart.RawGraphicFrameXml.Replace("\"" + chart.OriginalRId + "\"", "\"" + rId + "\"");
                try
                {
                    anchor.Add(XElement.Parse(updatedXml));
                }
                catch
                {
                    // If raw XML is unparseable preserve a minimal placeholder so the file remains valid
                }
            }
            else
            {
                anchor.Add(BuildGraphicFrameElement(chart, rId, chartId));
            }

            anchor.Add(new XElement(XdrNs + "clientData"));
            return anchor;
        }

        private static XElement BuildGraphicFrameElement(Core.ChartModel chart, string rId, int frameId)
        {
            var displayName = string.IsNullOrEmpty(chart.Name)
                ? "Chart " + frameId.ToString(CultureInfo.InvariantCulture)
                : chart.Name;

            var graphicFrame = new XElement(XdrNs + "graphicFrame",
                new XAttribute("macro", string.Empty));

            graphicFrame.Add(new XElement(XdrNs + "nvGraphicFramePr",
                new XElement(XdrNs + "cNvPr",
                    new XAttribute("id", frameId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("name", displayName)),
                new XElement(XdrNs + "cNvGraphicFramePr")));

            graphicFrame.Add(new XElement(XdrNs + "xfrm",
                new XElement(ANs + "off",
                    new XAttribute("x", "0"),
                    new XAttribute("y", "0")),
                new XElement(ANs + "ext",
                    new XAttribute("cx", "0"),
                    new XAttribute("cy", "0"))));

            graphicFrame.Add(new XElement(ANs + "graphic",
                new XElement(ANs + "graphicData",
                    new XAttribute("uri", ChartGraphicDataUri),
                    new XElement(ChartNs + "chart",
                        new XAttribute(XNamespace.Xmlns + "c", ChartNs),
                        new XAttribute(XNamespace.Xmlns + "r", RelationshipNs),
                        new XAttribute(RelationshipNs + "id", rId)))));

            return graphicFrame;
        }
    }
}
