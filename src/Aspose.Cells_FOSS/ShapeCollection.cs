using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents the collection of drawing objects (shapes) on a worksheet.
    /// </summary>
    public sealed class ShapeCollection
    {
        private readonly WorksheetModel _worksheetModel;

        internal ShapeCollection(WorksheetModel worksheetModel)
        {
            _worksheetModel = worksheetModel;
        }

        /// <summary>
        /// Gets the number of shapes in the collection.
        /// </summary>
        public int Count
        {
            get
            {
                return _worksheetModel.Shapes.Count;
            }
        }

        /// <summary>
        /// Gets the shape at the specified zero-based index.
        /// </summary>
        public Shape this[int index]
        {
            get
            {
                if (index < 0 || index >= _worksheetModel.Shapes.Count)
                {
                    throw new CellsException("Shape index " + index + " is out of range.");
                }

                return new Shape(_worksheetModel.Shapes[index]);
            }
        }

        /// <summary>
        /// Adds a new shape to the worksheet and returns its zero-based index.
        /// </summary>
        /// <param name="upperLeftRow">Zero-based row index of the upper-left anchor cell.</param>
        /// <param name="upperLeftColumn">Zero-based column index of the upper-left anchor cell.</param>
        /// <param name="lowerRightRow">Zero-based row index of the lower-right anchor cell.</param>
        /// <param name="lowerRightColumn">Zero-based column index of the lower-right anchor cell.</param>
        /// <param name="shapeType">The type of shape to add.</param>
        public int Add(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, AutoShapeType shapeType)
        {
            if (upperLeftRow < 0)
            {
                throw new CellsException("upperLeftRow must be non-negative.");
            }

            if (upperLeftColumn < 0)
            {
                throw new CellsException("upperLeftColumn must be non-negative.");
            }

            if (lowerRightRow < upperLeftRow)
            {
                throw new CellsException("lowerRightRow must be >= upperLeftRow.");
            }

            if (lowerRightColumn < upperLeftColumn)
            {
                throw new CellsException("lowerRightColumn must be >= upperLeftColumn.");
            }

            var index = _worksheetModel.Shapes.Count;
            var model = new ShapeModel
            {
                Name = "Shape " + (index + 1),
                UpperLeftRow = upperLeftRow,
                UpperLeftColumn = upperLeftColumn,
                LowerRightRow = lowerRightRow,
                LowerRightColumn = lowerRightColumn,
                GeometryType = Shape.AutoShapeTypeToGeometry(shapeType),
            };
            _worksheetModel.Shapes.Add(model);
            return index;
        }

        /// <summary>
        /// Removes the shape at the specified zero-based index.
        /// </summary>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _worksheetModel.Shapes.Count)
            {
                throw new CellsException("Shape index " + index + " is out of range.");
            }

            _worksheetModel.Shapes.RemoveAt(index);
        }
    }
}
