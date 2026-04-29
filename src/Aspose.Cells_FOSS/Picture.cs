using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a picture (image) anchored to a worksheet.
    /// </summary>
    public sealed class Picture
    {
        private readonly PictureModel _model;

        internal Picture(PictureModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets the display name of the picture.
        /// </summary>
        public string Name
        {
            get
            {
                return _model.Name;
            }
            set
            {
                _model.Name = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based row index of the upper-left anchor cell.
        /// </summary>
        public int UpperLeftRow
        {
            get
            {
                return _model.UpperLeftRow;
            }
            set
            {
                if (value < 0)
                {
                    throw new CellsException("UpperLeftRow must be non-negative.");
                }

                _model.UpperLeftRow = value;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based column index of the upper-left anchor cell.
        /// </summary>
        public int UpperLeftColumn
        {
            get
            {
                return _model.UpperLeftColumn;
            }
            set
            {
                if (value < 0)
                {
                    throw new CellsException("UpperLeftColumn must be non-negative.");
                }

                _model.UpperLeftColumn = value;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based row index of the lower-right anchor cell.
        /// </summary>
        public int LowerRightRow
        {
            get
            {
                return _model.LowerRightRow;
            }
            set
            {
                _model.LowerRightRow = value;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based column index of the lower-right anchor cell.
        /// </summary>
        public int LowerRightColumn
        {
            get
            {
                return _model.LowerRightColumn;
            }
            set
            {
                _model.LowerRightColumn = value;
            }
        }

        /// <summary>
        /// Gets the image format derived from the file extension.
        /// </summary>
        public ImageType ImageType
        {
            get
            {
                return ImageTypeFromExtension(_model.ImageExtension);
            }
        }

        /// <summary>
        /// Gets or sets the raw binary bytes of the image.
        /// Setting this also updates ImageType via the image bytes magic signature.
        /// </summary>
        public byte[] Data
        {
            get
            {
                return _model.ImageData;
            }
            set
            {
                if (value == null || value.Length == 0)
                {
                    throw new CellsException("Picture data must be non-null and non-empty.");
                }

                _model.ImageData = value;
                _model.ImageExtension = DetectExtension(value);
            }
        }

        internal static ImageType ImageTypeFromExtension(string extension)
        {
            switch (extension)
            {
                case "jpeg":
                case "jpg":
                    return ImageType.Jpeg;
                case "png":
                    return ImageType.Png;
                case "gif":
                    return ImageType.Gif;
                case "bmp":
                    return ImageType.Bmp;
                default:
                    return ImageType.Unknown;
            }
        }

        internal static string ExtensionFromImageType(ImageType type)
        {
            switch (type)
            {
                case ImageType.Jpeg: return "jpeg";
                case ImageType.Png: return "png";
                case ImageType.Gif: return "gif";
                case ImageType.Bmp: return "bmp";
                default: return "bin";
            }
        }

        internal static string DetectExtension(byte[] data)
        {
            if (data == null || data.Length < 4)
            {
                return "bin";
            }

            if (data[0] == 0xFF && data[1] == 0xD8)
            {
                return "jpeg";
            }

            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
            {
                return "png";
            }

            if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46)
            {
                return "gif";
            }

            if (data[0] == 0x42 && data[1] == 0x4D)
            {
                return "bmp";
            }

            return "bin";
        }

        internal static string ContentTypeFromExtension(string extension)
        {
            switch (extension)
            {
                case "jpeg":
                case "jpg":
                    return "image/jpeg";
                case "png":
                    return "image/png";
                case "gif":
                    return "image/gif";
                case "bmp":
                    return "image/bmp";
                default:
                    return "application/octet-stream";
            }
        }
    }
}
