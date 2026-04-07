using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.XlsxWorkbookArchiveHelpers;
using static Aspose.Cells_FOSS.XlsxWorkbookSerializerCommon;

namespace Aspose.Cells_FOSS;

internal static class XlsxDocumentProperties
{
    private static readonly XNamespace CorePropertiesNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DublinCoreNs = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DublinCoreTermsNs = "http://purl.org/dc/terms/";
    private static readonly XNamespace XsiNs = "http://www.w3.org/2001/XMLSchema-instance";
    private static readonly XNamespace ExtendedPropertiesNs = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    internal static XDocument? BuildCorePropertiesDocument(WorkbookModel model)
    {
        if (!model.DocumentProperties.Core.HasStoredState())
        {
            return null;
        }

        var core = model.DocumentProperties.Core;
        var root = new XElement(CorePropertiesNs + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", CorePropertiesNs),
            new XAttribute(XNamespace.Xmlns + "dc", DublinCoreNs),
            new XAttribute(XNamespace.Xmlns + "dcterms", DublinCoreTermsNs),
            new XAttribute(XNamespace.Xmlns + "xsi", XsiNs));

        AddStringElement(root, DublinCoreNs + "title", core.Title);
        AddStringElement(root, DublinCoreNs + "subject", core.Subject);
        AddStringElement(root, DublinCoreNs + "creator", core.Creator);
        AddStringElement(root, CorePropertiesNs + "keywords", core.Keywords);
        AddStringElement(root, DublinCoreNs + "description", core.Description);
        AddStringElement(root, CorePropertiesNs + "lastModifiedBy", core.LastModifiedBy);
        AddStringElement(root, CorePropertiesNs + "revision", core.Revision);
        AddStringElement(root, CorePropertiesNs + "category", core.Category);
        AddStringElement(root, CorePropertiesNs + "contentStatus", core.ContentStatus);
        AddDateElement(root, DublinCoreTermsNs + "created", core.Created);
        AddDateElement(root, DublinCoreTermsNs + "modified", core.Modified);

        return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
    }

    internal static XDocument? BuildExtendedPropertiesDocument(WorkbookModel model)
    {
        if (!model.DocumentProperties.Extended.HasStoredState())
        {
            return null;
        }

        var extended = model.DocumentProperties.Extended;
        var root = new XElement(ExtendedPropertiesNs + "Properties",
            new XAttribute(XNamespace.Xmlns + "vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"));

        AddStringElement(root, ExtendedPropertiesNs + "Application", extended.Application);
        AddStringElement(root, ExtendedPropertiesNs + "AppVersion", extended.AppVersion);
        AddStringElement(root, ExtendedPropertiesNs + "Company", extended.Company);
        AddStringElement(root, ExtendedPropertiesNs + "Manager", extended.Manager);
        AddIntElement(root, ExtendedPropertiesNs + "DocSecurity", extended.DocSecurity);
        AddStringElement(root, ExtendedPropertiesNs + "HyperlinkBase", extended.HyperlinkBase);
        AddBoolElement(root, ExtendedPropertiesNs + "ScaleCrop", extended.ScaleCrop);
        AddBoolElement(root, ExtendedPropertiesNs + "LinksUpToDate", extended.LinksUpToDate);
        AddBoolElement(root, ExtendedPropertiesNs + "SharedDoc", extended.SharedDoc);

        return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
    }

    internal static void LoadDocumentProperties(ZipArchive archive, WorkbookModel model, LoadDiagnostics diagnostics, LoadOptions options)
    {
        var relationshipTargets = LoadRootDocumentPropertiesRelationshipTargets(archive, diagnostics, options);
        var corePropertiesPartUri = string.Empty;
        if (relationshipTargets.TryGetValue(CorePropertiesRelationshipType, out var resolvedCorePropertiesPartUri)
            && !string.IsNullOrEmpty(resolvedCorePropertiesPartUri))
        {
            corePropertiesPartUri = resolvedCorePropertiesPartUri;
        }

        var extendedPropertiesPartUri = string.Empty;
        if (relationshipTargets.TryGetValue(ExtendedPropertiesRelationshipType, out var resolvedExtendedPropertiesPartUri)
            && !string.IsNullOrEmpty(resolvedExtendedPropertiesPartUri))
        {
            extendedPropertiesPartUri = resolvedExtendedPropertiesPartUri;
        }

        LoadCoreProperties(archive, model, diagnostics, options, corePropertiesPartUri);
        LoadExtendedProperties(archive, model, diagnostics, options, extendedPropertiesPartUri);
    }

    private static Dictionary<string, string> LoadRootDocumentPropertiesRelationshipTargets(ZipArchive archive, LoadDiagnostics diagnostics, LoadOptions options)
    {
        var relationshipTargets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var entry = GetEntry(archive, "/_rels/.rels");
        if (entry is null)
        {
            return relationshipTargets;
        }

        XDocument? document;
        try
        {
            document = LoadDocument(entry);
        }
        catch (Exception)
        {
            AddDocumentPropertiesIssue(diagnostics, options, "/_rels/.rels", "Package root relationships were malformed and document properties were ignored.");
            return relationshipTargets;
        }

        foreach (var relationship in document.Root?.Elements(PackageRelationshipNs + "Relationship") ?? Enumerable.Empty<XElement>())
        {
            var type = (string?)relationship.Attribute("Type");
            var target = (string?)relationship.Attribute("Target");
            var targetMode = (string?)relationship.Attribute("TargetMode");
            if (string.IsNullOrEmpty(type)
                || string.IsNullOrEmpty(target)
                || string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var relationshipType = type!;
            if (!string.Equals(relationshipType, CorePropertiesRelationshipType, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(relationshipType, ExtendedPropertiesRelationshipType, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!relationshipTargets.ContainsKey(relationshipType))
            {
                relationshipTargets[relationshipType] = ResolvePartUri("/", target!);
            }
        }

        return relationshipTargets;
    }

    private static void LoadCoreProperties(ZipArchive archive, WorkbookModel model, LoadDiagnostics diagnostics, LoadOptions options, string partUri)
    {
        if (string.IsNullOrEmpty(partUri))
        {
            return;
        }

        var entry = GetEntry(archive, partUri);
        if (entry is null)
        {
            return;
        }

        XDocument? document;
        try
        {
            document = LoadDocument(entry);
        }
        catch (Exception)
        {
            AddDocumentPropertiesIssue(diagnostics, options, partUri, "Core document properties part was malformed and was ignored.");
            return;
        }

        var root = document.Root;
        if (root is null)
        {
            return;
        }

        var core = model.DocumentProperties.Core;
        core.Title = ReadElementValue(root, DublinCoreNs + "title");
        core.Subject = ReadElementValue(root, DublinCoreNs + "subject");
        core.Creator = ReadElementValue(root, DublinCoreNs + "creator");
        core.Keywords = ReadElementValue(root, CorePropertiesNs + "keywords");
        core.Description = ReadElementValue(root, DublinCoreNs + "description");
        core.LastModifiedBy = ReadElementValue(root, CorePropertiesNs + "lastModifiedBy");
        core.Revision = ReadElementValue(root, CorePropertiesNs + "revision");
        core.Category = ReadElementValue(root, CorePropertiesNs + "category");
        core.ContentStatus = ReadElementValue(root, CorePropertiesNs + "contentStatus");
        core.Created = ReadDateElement(root, DublinCoreTermsNs + "created", diagnostics, options, partUri);
        core.Modified = ReadDateElement(root, DublinCoreTermsNs + "modified", diagnostics, options, partUri);
    }

    private static void LoadExtendedProperties(ZipArchive archive, WorkbookModel model, LoadDiagnostics diagnostics, LoadOptions options, string partUri)
    {
        if (string.IsNullOrEmpty(partUri))
        {
            return;
        }

        var entry = GetEntry(archive, partUri);
        if (entry is null)
        {
            return;
        }

        XDocument? document;
        try
        {
            document = LoadDocument(entry);
        }
        catch (Exception)
        {
            AddDocumentPropertiesIssue(diagnostics, options, partUri, "Extended document properties part was malformed and was ignored.");
            return;
        }

        var root = document.Root;
        if (root is null)
        {
            return;
        }

        var extended = model.DocumentProperties.Extended;
        extended.Application = ReadElementValue(root, ExtendedPropertiesNs + "Application");
        extended.AppVersion = ReadElementValue(root, ExtendedPropertiesNs + "AppVersion");
        extended.Company = ReadElementValue(root, ExtendedPropertiesNs + "Company");
        extended.Manager = ReadElementValue(root, ExtendedPropertiesNs + "Manager");
        extended.DocSecurity = ReadIntElement(root, ExtendedPropertiesNs + "DocSecurity", diagnostics, options, partUri);
        extended.HyperlinkBase = ReadElementValue(root, ExtendedPropertiesNs + "HyperlinkBase");
        extended.ScaleCrop = ReadBoolElement(root, ExtendedPropertiesNs + "ScaleCrop", diagnostics, options, partUri);
        extended.LinksUpToDate = ReadBoolElement(root, ExtendedPropertiesNs + "LinksUpToDate", diagnostics, options, partUri);
        extended.SharedDoc = ReadBoolElement(root, ExtendedPropertiesNs + "SharedDoc", diagnostics, options, partUri);
    }

    private static void AddStringElement(XElement parent, XName name, string value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            parent.Add(new XElement(name, value));
        }
    }

    private static void AddIntElement(XElement parent, XName name, int? value)
    {
        if (value.HasValue)
        {
            parent.Add(new XElement(name, value.Value.ToString(CultureInfo.InvariantCulture)));
        }
    }

    private static void AddBoolElement(XElement parent, XName name, bool? value)
    {
        if (value.HasValue)
        {
            parent.Add(new XElement(name, value.Value ? "true" : "false"));
        }
    }

    private static void AddDateElement(XElement parent, XName name, DateTime? value)
    {
        if (!value.HasValue)
        {
            return;
        }

        var element = new XElement(name, XmlConvert.ToString(value.Value, XmlDateTimeSerializationMode.RoundtripKind));
        element.SetAttributeValue(XsiNs + "type", "dcterms:W3CDTF");
        parent.Add(element);
    }

    private static string ReadElementValue(XElement parent, XName name)
    {
        return (parent.Element(name)?.Value ?? string.Empty).Trim();
    }

    private static DateTime? ReadDateElement(XElement parent, XName name, LoadDiagnostics diagnostics, LoadOptions options, string partUri)
    {
        var element = parent.Element(name);
        if (element is null)
        {
            return null;
        }

        try
        {
            return XmlConvert.ToDateTime(element.Value, XmlDateTimeSerializationMode.RoundtripKind);
        }
        catch (Exception)
        {
            AddDocumentPropertiesIssue(diagnostics, options, partUri, "Document property '" + name.LocalName + "' had an invalid timestamp and was ignored.");
            return null;
        }
    }

    private static int? ReadIntElement(XElement parent, XName name, LoadDiagnostics diagnostics, LoadOptions options, string partUri)
    {
        var element = parent.Element(name);
        if (element is null)
        {
            return null;
        }

        if (int.TryParse(element.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) && value >= 0)
        {
            return value;
        }

        AddDocumentPropertiesIssue(diagnostics, options, partUri, "Document property '" + name.LocalName + "' had an invalid integer value and was ignored.");
        return null;
    }

    private static bool? ReadBoolElement(XElement parent, XName name, LoadDiagnostics diagnostics, LoadOptions options, string partUri)
    {
        var element = parent.Element(name);
        if (element is null)
        {
            return null;
        }

        var rawValue = element.Value;
        if (rawValue == "1" || string.Equals(rawValue, "true", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (rawValue == "0" || string.Equals(rawValue, "false", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        AddDocumentPropertiesIssue(diagnostics, options, partUri, "Document property '" + name.LocalName + "' had an invalid Boolean value and was ignored.");
        return null;
    }

    private static void AddDocumentPropertiesIssue(LoadDiagnostics diagnostics, LoadOptions options, string partUri, string message)
    {
        AddIssue(diagnostics, options, new LoadIssue("WB-L004", DiagnosticSeverity.Warning, message)
        {
            PartUri = partUri,
        });
    }
}
