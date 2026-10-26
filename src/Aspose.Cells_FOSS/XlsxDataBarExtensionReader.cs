using System;
using System.Globalization;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    internal static class XlsxDataBarExtensionReader
    {
        private static readonly XNamespace MainNs =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace ExtensionNs =
            "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

        internal static void Apply(XElement worksheetRoot, XElement ruleElement,
            FormatConditionModel condition)
        {
            if (condition.Type != FormatConditionType.DataBar)
            {
                return;
            }

            string ruleId = null;
            foreach (var idElement in ruleElement.Descendants(ExtensionNs + "id"))
            {
                ruleId = idElement.Value;
                break;
            }

            var extensions = worksheetRoot.Element(MainNs + "extLst");
            if (string.IsNullOrEmpty(ruleId) || extensions == null)
            {
                return;
            }

            foreach (var extendedRule in extensions.Descendants(ExtensionNs + "cfRule"))
            {
                if (!string.Equals((string)extendedRule.Attribute("id"), ruleId,
                    StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var dataBar = extendedRule.Element(ExtensionNs + "dataBar");
                if (dataBar == null)
                {
                    return;
                }

                var gradient = (string)dataBar.Attribute("gradient");
                if (gradient != null)
                {
                    condition.DataBarGradient = gradient != "0"
                        && !string.Equals(gradient, "false", StringComparison.OrdinalIgnoreCase);
                }

                double minLength;
                double maxLength;
                if (TryReadLength(dataBar.Attribute("minLength"), out minLength))
                {
                    condition.DataBarMinLength = minLength;
                }

                if (TryReadLength(dataBar.Attribute("maxLength"), out maxLength))
                {
                    condition.DataBarMaxLength = maxLength;
                }

                if (condition.DataBarMinLength > condition.DataBarMaxLength)
                {
                    condition.DataBarMinLength = 0.10d;
                    condition.DataBarMaxLength = 0.90d;
                }

                foreach (var threshold in dataBar.Elements(ExtensionNs + "cfvo"))
                {
                    condition.DataBarAutoMinimum = string.Equals(
                        (string)threshold.Attribute("type"), "autoMin", StringComparison.OrdinalIgnoreCase);
                    break;
                }

                return;
            }
        }

        private static bool TryReadLength(XAttribute attribute, out double length)
        {
            length = 0d;
            double percent;
            if (attribute == null || !double.TryParse(attribute.Value, NumberStyles.Float,
                CultureInfo.InvariantCulture, out percent) || percent < 0d || percent > 100d)
            {
                return false;
            }

            length = percent / 100d;
            return true;
        }
    }
}
