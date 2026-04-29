using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a worksheet comment (legacy note) anchored to a single cell.
    /// </summary>
    public sealed class Comment
    {
        private readonly CommentModel _model;

        internal Comment(CommentModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets the zero-based row index of the annotated cell. Read-only after the comment is added.
        /// </summary>
        public int Row
        {
            get
            {
                return _model.Row;
            }
        }

        /// <summary>
        /// Gets the zero-based column index of the annotated cell. Read-only after the comment is added.
        /// </summary>
        public int Column
        {
            get
            {
                return _model.Column;
            }
        }

        /// <summary>
        /// Gets or sets the comment author name.
        /// </summary>
        public string Author
        {
            get
            {
                return _model.Author;
            }
            set
            {
                _model.Author = value == null ? string.Empty : value;
            }
        }

        /// <summary>
        /// Gets or sets the plain-text note content of the comment.
        /// </summary>
        public string Note
        {
            get
            {
                return _model.Note;
            }
            set
            {
                _model.Note = value == null ? string.Empty : value;
            }
        }

        /// <summary>
        /// Gets or sets whether the comment box is always shown (true) or hidden until hover (false).
        /// </summary>
        public bool IsVisible
        {
            get
            {
                return _model.IsVisible;
            }
            set
            {
                _model.IsVisible = value;
            }
        }

        /// <summary>
        /// Gets or sets the comment box width in pixels. Must be at least 1.
        /// </summary>
        public int Width
        {
            get
            {
                return _model.Width;
            }
            set
            {
                if (value < 1)
                {
                    throw new CellsException("Comment Width must be at least 1 pixel.");
                }

                _model.Width = value;
            }
        }

        /// <summary>
        /// Gets or sets the comment box height in pixels. Must be at least 1.
        /// </summary>
        public int Height
        {
            get
            {
                return _model.Height;
            }
            set
            {
                if (value < 1)
                {
                    throw new CellsException("Comment Height must be at least 1 pixel.");
                }

                _model.Height = value;
            }
        }
    }
}
