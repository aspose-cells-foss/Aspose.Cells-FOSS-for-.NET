using System;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>Reads portable page defaults from the public portion of a saved DEVMODEW.</summary>
    internal static class XlsxWorkbookPrinterSettings
    {
        internal static void Load(PageSetupModel setup, XElement root, ZipArchive archive,
            string worksheetUri, LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            var main = XlsxWorkbookSerializerCommon.MainNs;
            var rel = XlsxWorkbookSerializerCommon.RelationshipNs;
            var id = (string)root.Element(main + "pageSetup")?.Attribute(rel + "id");
            if (string.IsNullOrEmpty(id)
                || (setup.PaperSize != 0 && setup.Orientation != PageOrientation.Default)) return;
            var slash = worksheetUri.LastIndexOf('/');
            var relsUri = worksheetUri.Substring(0, slash + 1) + "_rels/" + worksheetUri.Substring(slash + 1) + ".rels";
            var relationships = XlsxWorkbookArchiveHelpers.LoadRelationships(archive, relsUri, worksheetUri);
            string target;
            if (!relationships.TryGetValue(id, out target))
            {
                Warn(diagnostics, options, sheetName);
                return;
            }
            var entry = XlsxWorkbookArchiveHelpers.GetEntry(archive, target);
            if (entry == null || entry.Length < 220)
            {
                Warn(diagnostics, options, sheetName);
                return;
            }
            var header = new byte[220];
            using (var stream = entry.Open())
            {
                var offset = 0;
                while (offset < header.Length)
                {
                    var read = stream.Read(header, offset, header.Length - offset);
                    if (read == 0) break;
                    offset += read;
                }
                if (offset != header.Length || !TryApply(header, entry.Length, setup))
                    Warn(diagnostics, options, sheetName);
            }
        }

        internal static bool TryApply(byte[] header, long length, PageSetupModel setup)
        {
            if (header == null || header.Length < 220) return false;
            var size = ReadUInt16(header, 68);
            var extra = ReadUInt16(header, 70);
            if (size < 220 || size + extra > length) return false;
            // dmFields marks which fields are meaningful. Never interpret driver-private bytes.
            var fields = ReadUInt16(header, 72);
            var paper = ReadUInt16(header, 78);
            if ((fields & 2) != 0 && paper > 0 && paper < 256 && setup.PaperSize == 0)
                setup.PaperSize = paper;
            var orientation = ReadUInt16(header, 76);
            if ((fields & 1) != 0 && setup.Orientation == PageOrientation.Default)
            {
                if (orientation == 1) setup.Orientation = PageOrientation.Portrait;
                if (orientation == 2) setup.Orientation = PageOrientation.Landscape;
            }
            return true;
        }

        private static int ReadUInt16(byte[] data, int offset)
        {
            return data[offset] | data[offset + 1] << 8;
        }

        private static void Warn(LoadDiagnostics diagnostics, LoadOptions options, string sheetName)
        {
            XlsxWorkbookArchiveHelpers.AddIssue(diagnostics, options,
                new LoadIssue("PS-L001", DiagnosticSeverity.Warning,
                    "Saved printer settings are missing or invalid; explicit page settings and default paper size were retained.", dataLossRisk: true)
                { SheetName = sheetName });
        }
    }
}
