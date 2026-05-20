using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents collection of pictures anchored to a worksheet.
    /// </summary>
    /// <remarks>
    /// Pictures allow embedding images directly into Excel worksheets from various sources.
    /// Supported image formats include JPEG, PNG, GIF, and BMP.
    /// Pictures can be loaded from files, streams, or byte arrays.
    /// Each picture can be independently positioned and sized using anchor coordinates.
    /// The collection provides methods to add, remove, and access pictures by index.
    /// </remarks>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// var worksheet = workbook.Worksheets[0];
    ///
    /// // Add pictures from different sources
    /// byte[] logoData = File.ReadAllBytes("logo.png");
    /// worksheet.Pictures.Add(2, 2, 5, 5, logoData);
    ///
    /// using (var stream = File.OpenRead("image.jpg"))
    /// {
    ///     worksheet.Pictures.Add(10, 2, 15, 10, stream);
    /// }
    ///
    /// Console.WriteLine($"Added {worksheet.Pictures.Count} pictures");
    /// </code>
    /// </example>
    public sealed class PictureCollection
    {
        private readonly WorksheetModel _worksheetModel;
        private readonly List<PictureModel> _models;

        internal PictureCollection(WorksheetModel worksheetModel)
        {
            _worksheetModel = worksheetModel;
            _models = worksheetModel.Pictures;
        }

        /// <summary>
        /// Gets the number of pictures on the worksheet.
        /// </summary>
        public int Count
        {
            get
            {
                return _models.Count;
            }
        }

        /// <summary>
        /// Gets the picture at the specified zero-based index.
        /// </summary>
        public Picture this[int index]
        {
            get
            {
                if (index < 0 || index >= _models.Count)
                {
                    throw new CellsException("Picture index " + index + " is out of range.");
                }

                return new Picture(_models[index]);
            }
        }

        /// <summary>
        /// Adds a picture to the worksheet from raw image bytes and returns its index.
        /// </summary>
        public int Add(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, byte[] imageData)
        {
            ValidateAnchor(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            if (imageData == null || imageData.Length == 0)
            {
                throw new CellsException("Image data must be non-null and non-empty.");
            }

            var model = CreateModel(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, imageData);
            _models.Add(model);
            return _models.Count - 1;
        }

        /// <summary>
        /// Adds a picture to the worksheet from a stream and returns its index.
        /// </summary>
        public int Add(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, Stream stream)
        {
            if (stream == null)
            {
                throw new CellsException("Image stream must not be null.");
            }

            byte[] data;
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                data = ms.ToArray();
            }

            return Add(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, data);
        }

        /// <summary>
        /// Adds a picture to the worksheet from a file path and returns its index.
        /// </summary>
        public int Add(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new CellsException("File name must be non-empty.");
            }

            var data = File.ReadAllBytes(fileName);
            var model = CreateModel(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, data);

            var ext = Path.GetExtension(fileName);
            if (!string.IsNullOrEmpty(ext) && ext.StartsWith(".", StringComparison.Ordinal))
            {
                ext = ext.Substring(1).ToLowerInvariant();
                if (ext == "jpg")
                {
                    ext = "jpeg";
                }

                model.ImageExtension = ext;
            }

            _models.Add(model);
            return _models.Count - 1;
        }

        /// <summary>
        /// Removes the picture at the specified zero-based index.
        /// </summary>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= _models.Count)
            {
                throw new CellsException("Picture index " + index + " is out of range.");
            }

            _models.RemoveAt(index);
        }

        private PictureModel CreateModel(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn, byte[] data)
        {
            ValidateAnchor(upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            var pictureNumber = _models.Count + 1;
            var model = new PictureModel
            {
                Name = "Picture " + pictureNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                UpperLeftRow = upperLeftRow,
                UpperLeftColumn = upperLeftColumn,
                LowerRightRow = lowerRightRow,
                LowerRightColumn = lowerRightColumn,
                ImageData = data,
                ImageExtension = Picture.DetectExtension(data),
            };
            return model;
        }

        private static void ValidateAnchor(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn)
        {
            if (upperLeftRow < 0)
            {
                throw new CellsException("UpperLeftRow must be non-negative.");
            }

            if (upperLeftColumn < 0)
            {
                throw new CellsException("UpperLeftColumn must be non-negative.");
            }

            if (lowerRightRow < upperLeftRow)
            {
                throw new CellsException("LowerRightRow must be greater than or equal to UpperLeftRow.");
            }

            if (lowerRightColumn < upperLeftColumn)
            {
                throw new CellsException("LowerRightColumn must be greater than or equal to UpperLeftColumn.");
            }
        }
    }
}
