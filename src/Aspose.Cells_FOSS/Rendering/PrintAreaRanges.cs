using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Resolves each print area separately without excluding chart source cells from layout.</summary>
    internal static class PrintAreaRanges
    {
        public static List<MergeRegion> Read(string printArea)
        {
            var ranges = new List<MergeRegion>();
            if (string.IsNullOrWhiteSpace(printArea)) return ranges;
            var normalized = XlsxWorkbookPageSetup.NormalizeLoadedPrintArea(printArea);
            foreach (var part in normalized.Split(','))
            {
                var ends = part.Replace("$", string.Empty).Split(':');
                var first = CellAddress.Parse(ends[0]);
                var last = CellAddress.Parse(ends[ends.Length - 1]);
                ranges.Add(new MergeRegion(first.RowIndex, first.ColumnIndex,
                    last.RowIndex - first.RowIndex + 1, last.ColumnIndex - first.ColumnIndex + 1));
            }
            return ranges;
        }
    }
}
