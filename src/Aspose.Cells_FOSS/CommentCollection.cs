using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents the collection of comments (legacy notes) on a worksheet.
    /// </summary>
    public sealed class CommentCollection
    {
        private readonly List<CommentModel> _models;

        internal CommentCollection(List<CommentModel> models)
        {
            _models = models;
        }

        /// <summary>
        /// Gets the number of comments on the worksheet.
        /// </summary>
        public int Count
        {
            get
            {
                return _models.Count;
            }
        }

        /// <summary>
        /// Gets the comment at the specified zero-based index.
        /// </summary>
        public Comment this[int index]
        {
            get
            {
                if (index < 0 || index >= _models.Count)
                {
                    throw new CellsException("Comment index " + index + " is out of range.");
                }

                return new Comment(_models[index]);
            }
        }

        /// <summary>
        /// Gets the comment at the given A1-style cell reference, or null if none exists.
        /// </summary>
        public Comment this[string cellName]
        {
            get
            {
                CellAddress address;
                if (!TryParseCell(cellName, out address))
                {
                    return null;
                }

                for (var i = 0; i < _models.Count; i++)
                {
                    if (_models[i].Row == address.RowIndex && _models[i].Column == address.ColumnIndex)
                    {
                        return new Comment(_models[i]);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Adds an empty comment at the given zero-based cell coordinates and returns it.
        /// </summary>
        public Comment Add(int row, int column)
        {
            if (row < 0)
            {
                throw new CellsException("Comment row must be non-negative.");
            }

            if (column < 0)
            {
                throw new CellsException("Comment column must be non-negative.");
            }

            for (var i = 0; i < _models.Count; i++)
            {
                if (_models[i].Row == row && _models[i].Column == column)
                {
                    throw new CellsException("A comment already exists at row " + row + ", column " + column + ".");
                }
            }

            var model = new CommentModel();
            model.Row = row;
            model.Column = column;
            _models.Add(model);
            return new Comment(model);
        }

        /// <summary>
        /// Adds an empty comment at the given A1-style cell reference and returns it.
        /// </summary>
        public Comment Add(string cellName)
        {
            CellAddress address;
            if (!TryParseCell(cellName, out address))
            {
                throw new CellsException("'" + cellName + "' is not a valid A1-style cell reference.");
            }

            return Add(address.RowIndex, address.ColumnIndex);
        }

        /// <summary>
        /// Removes the comment at the given zero-based index.
        /// </summary>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _models.Count)
            {
                throw new CellsException("Comment index " + index + " is out of range.");
            }

            _models.RemoveAt(index);
        }

        /// <summary>
        /// Removes the comment at the given A1-style cell reference. No-op if no comment exists there.
        /// </summary>
        public void RemoveAt(string cellName)
        {
            CellAddress address;
            if (!TryParseCell(cellName, out address))
            {
                return;
            }

            for (var i = 0; i < _models.Count; i++)
            {
                if (_models[i].Row == address.RowIndex && _models[i].Column == address.ColumnIndex)
                {
                    _models.RemoveAt(i);
                    return;
                }
            }
        }

        private static bool TryParseCell(string cellName, out CellAddress address)
        {
            try
            {
                address = CellAddress.Parse(cellName);
                return true;
            }
            catch (ArgumentException)
            {
                address = default(CellAddress);
                return false;
            }
        }
    }
}
