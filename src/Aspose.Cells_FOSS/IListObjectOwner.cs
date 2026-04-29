using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Allows ListObject to call back into its owning collection for cross-table validation.
    /// </summary>
    internal interface IListObjectOwner
    {
        void ValidateUniqueDisplayName(string displayName, ListObjectModel skipModel);
        void ValidateNoOverlap(int startRow, int startColumn, int endRow, int endColumn, ListObjectModel skipModel);
        void RemoveTable(ListObjectModel model);
    }
}
