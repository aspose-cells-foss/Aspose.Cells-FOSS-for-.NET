using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents one shared-string-table entry, including optional rich-text runs.
    /// </summary>
    internal sealed class SharedStringEntry
    {
        /// <summary>
        /// Gets or sets the flattened text.
        /// </summary>
        public string Text { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the formatting runs.
        /// </summary>
        public List<RichTextRunValue> Runs { get; set; }

        internal SharedStringEntry Clone()
        {
            var clone = new SharedStringEntry();
            clone.Text = Text ?? string.Empty;
            if (Runs != null && Runs.Count > 0)
            {
                clone.Runs = new List<RichTextRunValue>(Runs.Count);
                for (var index = 0; index < Runs.Count; index++)
                {
                    clone.Runs.Add(Runs[index].Clone());
                }
            }

            return clone;
        }
    }
}
