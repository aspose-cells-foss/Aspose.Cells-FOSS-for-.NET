using System;
using System.Collections.Generic;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfFontUsageRegistry : IDisposable
    {
        private readonly List<PdfType3FontUsage> _usages = new List<PdfType3FontUsage>();

        public IList<PdfType3FontUsage> Usages
        {
            get { return _usages; }
        }

        public void AddRuns(IList<PdfTextRunRecord> runs)
        {
            for (var i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                var usage = UsageFor(run.Typeface, run.UsePositiveYFontMatrix);
                usage.AddText(run.Text);
            }
        }

        public void FinalizeSubsets()
        {
            var resourceIndex = 1;
            for (var i = 0; i < _usages.Count; i++)
            {
                _usages[i].FinalizeSubsets(ref resourceIndex);
            }
        }

        public PdfType3FontUsage UsageFor(SKTypeface typeface, bool usePositiveYFontMatrix)
        {
            for (var i = 0; i < _usages.Count; i++)
            {
                if (ReferenceEquals(_usages[i].Typeface, typeface)
                    && _usages[i].UsePositiveYFontMatrix == usePositiveYFontMatrix)
                {
                    return _usages[i];
                }
            }

            var usage = new PdfType3FontUsage(typeface, usePositiveYFontMatrix);
            _usages.Add(usage);
            return usage;
        }

        public void Dispose()
        {
            for (var i = 0; i < _usages.Count; i++)
            {
                _usages[i].Dispose();
            }

            _usages.Clear();
        }
    }
}
