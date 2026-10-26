using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;
using static Aspose.Cells_FOSS.XlsxWorkbookStylesValueHelpers;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxStyleAlignmentWriter
    {
        internal static XElement Build(AlignmentValue alignment)
        {
            var element = new XElement(MainNs + "alignment");
            var hasValue = false;
            var horizontal = GetHorizontalAlignmentName(alignment.Horizontal);
            if (!string.IsNullOrEmpty(horizontal))
            {
                element.SetAttributeValue("horizontal", horizontal);
                hasValue = true;
            }
            var vertical = GetVerticalAlignmentName(alignment.Vertical);
            if (!string.IsNullOrEmpty(vertical))
            {
                element.SetAttributeValue("vertical", vertical);
                hasValue = true;
            }
            if (alignment.WrapText)
            {
                element.SetAttributeValue("wrapText", 1);
                hasValue = true;
            }
            if (alignment.IndentLevel > 0)
            {
                element.SetAttributeValue("indent", alignment.IndentLevel);
                hasValue = true;
            }
            if (alignment.TextRotation != 0)
            {
                element.SetAttributeValue("textRotation", alignment.TextRotation);
                hasValue = true;
            }
            if (alignment.ShrinkToFit)
            {
                element.SetAttributeValue("shrinkToFit", 1);
                hasValue = true;
            }
            if (alignment.ReadingOrder != 0)
            {
                element.SetAttributeValue("readingOrder", alignment.ReadingOrder);
                hasValue = true;
            }
            if (alignment.RelativeIndent != 0)
            {
                element.SetAttributeValue("relativeIndent", alignment.RelativeIndent);
                hasValue = true;
            }
            return hasValue ? element : null;
        }
    }
}
