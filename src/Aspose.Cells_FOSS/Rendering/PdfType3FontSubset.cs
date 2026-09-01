using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfType3FontSubset
    {
        public string ResourceName;
        public bool UsePositiveYFontMatrix;
        public readonly List<string> Tokens = new List<string>();
        public readonly Dictionary<string, byte> TokenCodes = new Dictionary<string, byte>();
        public int FontObjectId;
        public int ToUnicodeObjectId;
    }
}
