using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

internal static class XlsxWorkbookArchiveHelpers
{
    internal static Stream EnsureSeekable(Stream stream)
    {
        if (stream.CanSeek)
        {
            return stream;
        }

        var memory = new MemoryStream();
        stream.CopyTo(memory);
        memory.Position = 0;
        return memory;
    }

    internal static ZipArchiveEntry? GetEntry(ZipArchive archive, string uri)
    {
        return archive.GetEntry(uri.TrimStart('/').Replace('\\', '/'));
    }

    internal static XDocument LoadDocument(ZipArchiveEntry entry)
    {
        using var rawStream = entry.Open();
        using var stream = EnsureSeekable(rawStream);
        stream.Position = 0;
        return XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace);
    }

    internal static Dictionary<string, string> LoadRelationships(ZipArchive archive, string relationshipsUri, string sourceUri)
    {
        var entry = GetEntry(archive, relationshipsUri);
        if (entry is null)
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        var document = LoadDocument(entry);
        var relationships = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var relationship in document.Root?.Elements(XlsxWorkbookSerializerCommon.PackageRelationshipNs + "Relationship") ?? Enumerable.Empty<XElement>())
        {
            var id = (string?)relationship.Attribute("Id");
            var target = (string?)relationship.Attribute("Target");
            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target))
            {
                continue;
            }

            var relationshipId = id!;
            relationships[relationshipId] = ResolvePartUri(sourceUri, target!);
        }

        return relationships;
    }

    internal static string FindRelationshipTarget(IReadOnlyDictionary<string, string> workbookRelationships, string expectedPath)
    {
        foreach (var pair in workbookRelationships)
        {
            if (RelationshipPointsTo(pair.Value, expectedPath))
            {
                return pair.Value;
            }
        }

        return string.Empty;
    }

    internal static string FindRelationshipTargetByType(IReadOnlyDictionary<string, string> workbookRelationships, string relationshipType)
    {
        foreach (var pair in workbookRelationships)
        {
            if (GetRelationshipTargetKind(pair.Value) == relationshipType)
            {
                return pair.Value;
            }
        }

        return string.Empty;
    }

    internal static IReadOnlyList<string> LoadSharedStrings(ZipArchive archive, IReadOnlyDictionary<string, string> workbookRelationships, LoadOptions options, LoadDiagnostics diagnostics)
    {
        var sharedStringsUri = FindRelationshipTarget(workbookRelationships, "/xl/sharedStrings.xml");
        if (string.IsNullOrEmpty(sharedStringsUri))
        {
            if (workbookRelationships.Count == 0)
            {
                sharedStringsUri = "/xl/sharedStrings.xml";
            }
            else
            {
                sharedStringsUri = FindRelationshipTargetByType(workbookRelationships, XlsxWorkbookSerializerCommon.SharedStringsRelationshipType);
            }
        }

        if (string.IsNullOrEmpty(sharedStringsUri))
        {
            sharedStringsUri = "/xl/sharedStrings.xml";
        }

        var entry = GetEntry(archive, sharedStringsUri);
        if (entry is null)
        {
            return Array.Empty<string>();
        }

        var document = LoadDocument(entry);
        var items = new List<string>();
        foreach (var item in document.Root?.Elements(XlsxWorkbookSerializerCommon.MainNs + "si") ?? Enumerable.Empty<XElement>())
        {
            items.Add(ReadInlineString(item));
        }

        return items;
    }

    internal static HashSet<int> LoadDateStyleIndexes(ZipArchive archive, IReadOnlyDictionary<string, string> workbookRelationships, LoadOptions options, LoadDiagnostics diagnostics)
    {
        var stylesUri = FindRelationshipTarget(workbookRelationships, "/xl/styles.xml");
        if (string.IsNullOrEmpty(stylesUri))
        {
            stylesUri = "/xl/styles.xml";
        }

        var entry = GetEntry(archive, stylesUri);
        if (entry is null)
        {
            return new HashSet<int>();
        }

        var document = LoadDocument(entry);
        var customFormats = new Dictionary<int, string>();
        foreach (var numFmt in document.Root?.Element(XlsxWorkbookSerializerCommon.MainNs + "numFmts")?.Elements(XlsxWorkbookSerializerCommon.MainNs + "numFmt") ?? Enumerable.Empty<XElement>())
        {
            var id = ParseIntAttribute(numFmt.Attribute("numFmtId"));
            var code = (string?)numFmt.Attribute("formatCode");
            if (id.HasValue && !string.IsNullOrEmpty(code))
            {
                customFormats[id.Value] = code!;
            }
        }

        var dateStyleIndexes = new HashSet<int>();
        var cellXfs = document.Root?.Element(XlsxWorkbookSerializerCommon.MainNs + "cellXfs")?.Elements(XlsxWorkbookSerializerCommon.MainNs + "xf").ToList() ?? new List<XElement>();
        for (var index = 0; index < cellXfs.Count; index++)
        {
            var numFmtId = ParseIntAttribute(cellXfs[index].Attribute("numFmtId"));
            if (!numFmtId.HasValue)
            {
                continue;
            }

            if (XlsxWorkbookSerializerCommon.BuiltInDateFormats.Contains(numFmtId.Value))
            {
                dateStyleIndexes.Add(index);
                continue;
            }

            if (customFormats.TryGetValue(numFmtId.Value, out var formatCode) && LooksLikeDateFormat(formatCode))
            {
                dateStyleIndexes.Add(index);
            }
        }

        return dateStyleIndexes;
    }

    internal static bool ParseBoolAttribute(XAttribute? attribute)
    {
        if (attribute is null)
        {
            return false;
        }

        var value = attribute.Value;
        return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    internal static int? ParseIntAttribute(XAttribute? attribute)
    {
        if (attribute is null)
        {
            return null;
        }

        return int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : null;
    }

    internal static double? ParseDoubleAttribute(XAttribute? attribute)
    {
        if (attribute is null)
        {
            return null;
        }

        return double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? value : null;
    }

    internal static string NormalizeFormula(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalizedValue = value!.Trim();
        return normalizedValue.StartsWith("=", StringComparison.Ordinal) ? normalizedValue.Substring(1) : normalizedValue;
    }

    internal static string ReadInlineString(XElement? inlineStringElement)
    {
        if (inlineStringElement is null)
        {
            return string.Empty;
        }

        var textNodes = new List<string>();
        foreach (var element in inlineStringElement.DescendantsAndSelf())
        {
            if (element.Name == XlsxWorkbookSerializerCommon.MainNs + "t")
            {
                textNodes.Add(element.Value);
            }
        }

        return textNodes.Count > 0 ? string.Concat(textNodes) : inlineStringElement.Value;
    }

    internal static bool TryParseNumber(string rawValue, out object? numberValue)
    {
        if (!rawValue.Contains('e') && !rawValue.Contains('E') && decimal.TryParse(rawValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue))
        {
            if (!rawValue.Contains('.') && !rawValue.Contains(',') && decimal.Truncate(decimalValue) == decimalValue && decimalValue >= int.MinValue && decimalValue <= int.MaxValue)
            {
                numberValue = (int)decimalValue;
                return true;
            }

            numberValue = decimalValue;
            return true;
        }

        if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var doubleValue))
        {
            if (Math.Abs(doubleValue % 1d) < double.Epsilon && doubleValue >= int.MinValue && doubleValue <= int.MaxValue)
            {
                numberValue = (int)doubleValue;
                return true;
            }

            numberValue = doubleValue;
            return true;
        }

        numberValue = null;
        return false;
    }

    internal static void AddIssue(LoadDiagnostics diagnostics, LoadOptions options, LoadIssue issue)
    {
        diagnostics.Add(issue);
        options.WarningCallback?.Warning(new WarningInfo
        {
            Code = issue.Code,
            Severity = issue.Severity,
            Message = issue.Message,
            DataLossRisk = issue.DataLossRisk,
            PartUri = issue.PartUri,
            SheetName = issue.SheetName,
            CellRef = issue.CellRef,
            RowIndex = issue.RowIndex,
        });
    }

    internal static string ResolveWorksheetUri(int sheetIndex, string? relationshipId, IReadOnlyDictionary<string, string> workbookRelationships, ZipArchive archive, LoadDiagnostics diagnostics, string sheetName, LoadOptions options)
    {
        var relationshipKey = relationshipId;
        if (!string.IsNullOrWhiteSpace(relationshipKey) && workbookRelationships.TryGetValue(relationshipKey!, out var worksheetUri))
        {
            return worksheetUri;
        }

        var fallbackUri = $"/xl/worksheets/sheet{sheetIndex + 1}.xml";
        if (GetEntry(archive, fallbackUri) is not null)
        {
            AddIssue(diagnostics, options, new LoadIssue("PKG-R001", DiagnosticSeverity.Recoverable, "Worksheet relationship metadata was incomplete; the worksheet part was resolved by convention.", repairApplied: true)
            {
                SheetName = sheetName,
                PartUri = fallbackUri,
            });
        }

        return fallbackUri;
    }

    internal static string ResolvePartUri(string sourceUri, string target)
    {
        if (target.StartsWith("/", StringComparison.Ordinal))
        {
            return target;
        }

        var baseUri = new Uri($"http://package{sourceUri}", UriKind.Absolute);
        var resolvedUri = new Uri(baseUri, target);
        return resolvedUri.AbsolutePath;
    }

    internal static bool RelationshipPointsTo(string candidate, string expectedUri)
    {
        return string.Equals(candidate, expectedUri, StringComparison.OrdinalIgnoreCase);
    }

    internal static string GetRelationshipTargetKind(string candidate)
    {
        if (candidate.EndsWith("/sharedStrings.xml", StringComparison.OrdinalIgnoreCase))
        {
            return XlsxWorkbookSerializerCommon.SharedStringsRelationshipType;
        }

        if (candidate.EndsWith("/styles.xml", StringComparison.OrdinalIgnoreCase))
        {
            return XlsxWorkbookSerializerCommon.StylesRelationshipType;
        }

        if (candidate.IndexOf("/worksheets/", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return XlsxWorkbookSerializerCommon.WorksheetRelationshipType;
        }

        return string.Empty;
    }

    internal static bool LooksLikeDateFormat(string formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode))
        {
            return false;
        }

        if (formatCode.IndexOf("[$-F400]", StringComparison.OrdinalIgnoreCase) >= 0
            || formatCode.IndexOf("[$-F800]", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return true;
        }

        var builder = new StringBuilder(formatCode.Length);
        var inQuote = false;
        var inBracket = false;

        for (var index = 0; index < formatCode.Length; index++)
        {
            var character = formatCode[index];
            if (character == '"')
            {
                inQuote = !inQuote;
                continue;
            }

            if (inQuote)
            {
                continue;
            }

            if (character == '[')
            {
                inBracket = true;
                continue;
            }

            if (character == ']' && inBracket)
            {
                inBracket = false;
                continue;
            }

            if (inBracket)
            {
                builder.Append(char.ToLowerInvariant(character));
                continue;
            }

            if (character == '\\' || character == '_')
            {
                index++;
                continue;
            }

            if (character == '*')
            {
                continue;
            }

            builder.Append(char.ToLowerInvariant(character));
        }

        var normalized = builder.ToString();
        return normalized.IndexOf("yy", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("yyyy", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("dd", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("ddd", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("mm", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("mmm", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("m/", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("/m", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("h", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("am/pm", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("a/p", StringComparison.Ordinal) >= 0
            || normalized.IndexOf("ss", StringComparison.Ordinal) >= 0;
    }
}



