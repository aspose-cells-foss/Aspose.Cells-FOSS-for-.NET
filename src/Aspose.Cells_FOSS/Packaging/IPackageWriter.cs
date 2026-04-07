namespace Aspose.Cells_FOSS.Packaging;

public interface IPackageWriter
{
    void Write(Stream stream, PackageModel packageModel);
}