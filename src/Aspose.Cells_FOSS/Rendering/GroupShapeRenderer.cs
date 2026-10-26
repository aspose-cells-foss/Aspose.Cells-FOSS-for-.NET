using System;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Positions a DrawingML group's child shapes in its anchor rectangle.</summary>
    internal sealed class GroupShapeRenderer
    {
        private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private readonly RenderContext _context;

        public GroupShapeRenderer(RenderContext context)
        {
            _context = context;
        }

        public bool TryDraw(SKCanvas canvas, PageLayout page, SKRect frame, string xml)
        {
            if (string.IsNullOrEmpty(xml) || xml.IndexOf("grpSp", StringComparison.Ordinal) < 0)
            {
                return false;
            }

            XElement root;
            try
            {
                root = XElement.Parse(xml);
            }
            catch (Exception)
            {
                return false;
            }

            if (root.Name != Xdr + "grpSp")
            {
                return false;
            }

            DrawGroup(canvas, page, frame, root);
            return true;
        }

        private void DrawGroup(SKCanvas canvas, PageLayout page, SKRect frame, XElement root)
        {
            var properties = root.Element(Xdr + "grpSpPr");
            var transform = properties != null ? properties.Element(A + "xfrm") : null;
            var childOffset = transform != null ? transform.Element(A + "chOff") : null;
            var childExtent = transform != null ? transform.Element(A + "chExt") : null;
            if (childOffset == null || childExtent == null)
            {
                return;
            }

            var originX = Number(childOffset, "x");
            var originY = Number(childOffset, "y");
            var extentX = Number(childExtent, "cx");
            var extentY = Number(childExtent, "cy");
            if (extentX <= 0d || extentY <= 0d)
            {
                return;
            }

            var renderer = new ShapeRenderer(_context);
            foreach (var child in root.Elements())
            {
                if (child.Name != Xdr + "sp" && child.Name != Xdr + "pic" && child.Name != Xdr + "grpSp")
                {
                    continue;
                }

                var shapeProperties = child.Element(child.Name == Xdr + "grpSp" ? Xdr + "grpSpPr" : Xdr + "spPr");
                var childTransform = shapeProperties != null ? shapeProperties.Element(A + "xfrm") : null;
                var offset = childTransform != null ? childTransform.Element(A + "off") : null;
                var extent = childTransform != null ? childTransform.Element(A + "ext") : null;
                if (offset == null || extent == null)
                {
                    continue;
                }

                var left = frame.Left + (float)((Number(offset, "x") - originX) / extentX * frame.Width);
                var top = frame.Top + (float)((Number(offset, "y") - originY) / extentY * frame.Height);
                var width = (float)(Number(extent, "cx") / extentX * frame.Width);
                var height = (float)(Number(extent, "cy") / extentY * frame.Height);
                if (width <= 0f || height <= 0f)
                {
                    continue;
                }

                var childFrame = new SKRect(left, top, left + width, top + height);
                if (child.Name == Xdr + "grpSp")
                {
                    DrawGroup(canvas, page, childFrame, child);
                }
                else if (child.Name == Xdr + "pic")
                {
                    DrawPicture(canvas, page, childFrame, child);
                }
                else
                {
                    var parsed = ShapeXmlParser.Parse(child.ToString(), "rect", _context.Colors);
                    if (parsed != null)
                    {
                        if (parsed.TextWrap && !string.IsNullOrEmpty(parsed.Text))
                        {
                            // Excel's grouped text boxes use a wider effective font than the
                            // fallback Calibri metrics, which also changes their line breaks.
                            parsed.TextSizePt *= 1.18d;
                        }

                        renderer.Draw(canvas, page, childFrame, parsed);
                    }
                }
            }
        }

        private void DrawPicture(SKCanvas canvas, PageLayout page, SKRect frame, XElement picture)
        {
            var blipFill = picture.Element(Xdr + "blipFill");
            var blip = blipFill != null ? blipFill.Element(A + "blip") : null;
            var relationshipId = blip != null ? (string)blip.Attribute(R + "embed") : null;
            if (string.IsNullOrEmpty(relationshipId) || page == null)
            {
                return;
            }

            ShapeImageModel source = null;
            foreach (var image in page.Sheet.Sheet.ShapeImages)
            {
                if (string.Equals(image.OriginalRId, relationshipId, StringComparison.Ordinal))
                {
                    source = image;
                    break;
                }
            }

            if (source == null)
            {
                return;
            }

            var imageModel = new PictureModel { ImageData = source.ImageData, ImageExtension = source.Extension };
            var decoded = _context.Pictures.GetImage(imageModel, frame.Width, frame.Height);
            if (decoded == null)
            {
                return;
            }

            var imageFrame = frame;
            var crop = blipFill.Element(A + "srcRect");
            if (crop != null)
            {
                var left = Number(crop, "l") / 100000d;
                var top = Number(crop, "t") / 100000d;
                var right = Number(crop, "r") / 100000d;
                var bottom = Number(crop, "b") / 100000d;
                if (left + right < 1d && top + bottom < 1d)
                {
                    imageFrame = new SKRect(
                        frame.Left - (float)(frame.Width * left / (1d - left - right)),
                        frame.Top - (float)(frame.Height * top / (1d - top - bottom)),
                        frame.Right + (float)(frame.Width * right / (1d - left - right)),
                        frame.Bottom + (float)(frame.Height * bottom / (1d - top - bottom)));
                }
            }

            using (var paint = new SKPaint { IsAntialias = true, FilterQuality = SKFilterQuality.None })
            {
                canvas.Save();
                canvas.ClipRect(frame);
                canvas.DrawImage(decoded, imageFrame, paint);
                canvas.Restore();
            }
        }

        private static double Number(XElement element, string attribute)
        {
            double result;
            var text = (string)element.Attribute(attribute);
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
                ? result : 0d;
        }
    }
}
