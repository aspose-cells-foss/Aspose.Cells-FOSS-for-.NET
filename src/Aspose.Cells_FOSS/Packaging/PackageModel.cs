namespace Aspose.Cells_FOSS.Packaging;

public sealed class PackageModel
{
    public List<PackagePartDescriptor> Parts { get; } = new List<PackagePartDescriptor>();
    public List<RelationshipDescriptor> Relationships { get; } = new List<RelationshipDescriptor>();
    public Dictionary<string, byte[]> UnsupportedParts { get; } = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
}
