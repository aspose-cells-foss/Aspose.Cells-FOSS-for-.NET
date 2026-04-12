namespace Aspose.Cells_FOSS.Packaging;

/// <summary>
/// Defines a writer for package models.
/// </summary>
public interface IPackageWriter
{
    /// <summary>
    /// Writes the specified package model to the target stream.
    /// </summary>
    /// <param name="stream">The stream.</param>
    /// <param name="packageModel">The package model.</param>
    void Write(Stream stream, PackageModel packageModel);
}
