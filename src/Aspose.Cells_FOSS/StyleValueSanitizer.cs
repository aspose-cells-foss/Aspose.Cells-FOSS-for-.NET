using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS
{
    internal static class StyleValueSanitizer
    {
        internal static int NormalizeIndentLevel(int? value)
        {
            if (!value.HasValue || value.Value < 0 || value.Value > 250)
            {
                return 0;
            }

            return value.Value;
        }

        internal static int NormalizeTextRotation(int? value)
        {
            if (!value.HasValue)
            {
                return 0;
            }

            if (value.Value == 255)
            {
                return 255;
            }

            if (value.Value < 0 || value.Value > 180)
            {
                return 0;
            }

            return value.Value;
        }

        internal static int NormalizeReadingOrder(int? value)
        {
            if (!value.HasValue || value.Value < 0 || value.Value > 2)
            {
                return 0;
            }

            return value.Value;
        }
    }
}
