namespace Aspose.Cells_FOSS.Packaging;

public sealed class PackageLoadContext
{
    public PackageLoadContext(object? workbook, PackageModel package)
    {
        Workbook = workbook;
        Package = package;
    }

    public object? Workbook { get; }
    public PackageModel Package { get; }
}
