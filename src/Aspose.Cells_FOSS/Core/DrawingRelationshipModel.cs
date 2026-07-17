namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents a preserved relationship from a worksheet drawing part that is referenced
    /// by raw shape XML and must survive round-trip save.
    /// </summary>
    internal sealed class DrawingRelationshipModel
    {
        /// <summary>
        /// Gets or sets the original relationship identifier from the source drawing rels part.
        /// </summary>
        public string OriginalRId { get; set; }

        /// <summary>
        /// Gets or sets the relationship type URI.
        /// </summary>
        public string RelationshipType { get; set; }

        /// <summary>
        /// Gets or sets the original relative target as written in the drawing rels part.
        /// </summary>
        public string Target { get; set; }

        /// <summary>
        /// Gets or sets the optional TargetMode attribute value.
        /// </summary>
        public string TargetMode { get; set; }
    }
}
