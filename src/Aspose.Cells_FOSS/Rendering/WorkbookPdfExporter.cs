using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Entry point for XLSX-to-PDF export. Orchestrates the layout, pagination, and SkiaSharp
    /// rendering passes for every visible worksheet and streams the result as a single PDF document.
    /// </summary>
    internal static class WorkbookPdfExporter
    {
        public static void Export(WorkbookModel workbook, Stream stream, PdfSaveOptions options)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (options == null) options = new PdfSaveOptions();

            using (var context = new RenderContext(workbook, options.DefaultFont))
            {
                if (options.UseType3TextOptimization)
                {
                    context.PdfTextSession = new PdfTextDocumentSession();
                    context.EnableWorksheetTextOptimization = true;
                }

                var pages = BuildPages(context, workbook, options);

                for (var i = 0; i < pages.Count; i++)
                {
                    pages[i].PageNumber = i + 1;
                    pages[i].TotalPages = pages.Count;
                }

                var metadata = new SKDocumentPdfMetadata
                {
                    Creator = "Aspose.Cells FOSS",
                    Producer = "Aspose.Cells FOSS (SkiaSharp)",
                };

                using (var pdfBuffer = new MemoryStream())
                using (var wstream = new SKManagedWStream(pdfBuffer))
                using (var document = SKDocument.CreatePdf(wstream, metadata))
                {
                    var renderer = new SheetPdfRenderer(context);

                    if (pages.Count == 0)
                    {
                        // A valid PDF needs at least one page; emit a blank default (A4) page.
                        double w, h;
                        PaperSizes.GetDimensionsPoints(0, out w, out h);
                        var canvas = document.BeginPage((float)w, (float)h);
                        document.EndPage();
                    }

                    foreach (var page in pages)
                    {
                        var canvas = document.BeginPage((float)page.PageWidthPt, (float)page.PageHeightPt);
                        RenderPdfPage(canvas, renderer, page);
                        document.EndPage();
                    }

                    document.Close();

                    var bytes = pdfBuffer.ToArray();
                    bytes = new PdfImageObjectWriter().Rewrite(bytes, workbook);
                    if (context.PdfTextSession != null && context.PdfTextSession.HasRuns)
                    {
                        bytes = new PdfTextObjectWriter(options.UseTrueTypeSubsetPrototype).Rewrite(bytes, context.PdfTextSession);
                    }

                    bytes = new PdfContentStreamOptimizer().Rewrite(bytes);

                    stream.Write(bytes, 0, bytes.Length);
                }
            }
        }

        private static List<PageLayout> BuildPages(RenderContext context, WorkbookModel workbook, PdfSaveOptions options)
        {
            var pages = new List<PageLayout>();

            for (var sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                var sheet = workbook.Worksheets[sheetIndex];
                if (sheet.Visibility != SheetVisibility.Visible)
                {
                    continue;
                }

                var layout = SheetLayout.Build(context, sheet);
                if (layout.IsEmpty)
                {
                    continue;
                }

                pages.AddRange(Paginator.Paginate(layout, sheetIndex, options));
            }

            return pages;
        }

        private static void RenderPdfPage(SKCanvas canvas, SheetPdfRenderer renderer, PageLayout page)
        {
            renderer.RenderPage(canvas, page);
        }
    }
}
