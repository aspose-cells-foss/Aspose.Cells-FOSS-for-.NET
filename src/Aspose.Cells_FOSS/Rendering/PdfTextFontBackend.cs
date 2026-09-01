using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Rendering
{
    internal interface PdfTextFontBackend : IDisposable
    {
        void Initialize(IList<PdfTextRunRecord> runs);

        IDictionary<int, byte[]> BuildObjects(ref int nextObjectId);

        string BuildFontResourceText();

        IList<PdfEncodedTextSegment> EncodeRun(PdfTextRunRecord run);
    }
}
