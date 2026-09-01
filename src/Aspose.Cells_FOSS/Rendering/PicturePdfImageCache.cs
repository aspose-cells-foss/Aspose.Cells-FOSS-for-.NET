using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Reuses decoded picture resources across pages so the PDF backend can reference the same image
    /// object instead of re-materializing identical content for every draw call.
    /// </summary>
    internal sealed class PicturePdfImageCache : IDisposable
    {
        private readonly Dictionary<string, SKImage> _imagesByPictureVariant = new Dictionary<string, SKImage>(StringComparer.Ordinal);
        private readonly Dictionary<string, SKData> _dataByKey = new Dictionary<string, SKData>(StringComparer.Ordinal);
        private readonly Dictionary<string, SKImage> _imagesByKey = new Dictionary<string, SKImage>(StringComparer.Ordinal);

        public SKImage GetImage(PictureModel picture, float widthPt, float heightPt)
        {
            if (picture == null || picture.ImageData == null || picture.ImageData.Length == 0)
            {
                return null;
            }

            var targetWidthPx = ToTargetPixels(widthPt);
            var targetHeightPx = ToTargetPixels(heightPt);
            var variantKey = BuildVariantKey(picture.ImageData, targetWidthPx, targetHeightPx);

            SKImage cached;
            if (_imagesByPictureVariant.TryGetValue(variantKey, out cached))
            {
                return cached;
            }

            var key = BuildKey(picture.ImageData);

            SKData data;
            if (!_dataByKey.TryGetValue(key, out data))
            {
                data = SKData.CreateCopy(picture.ImageData);
                _dataByKey[key] = data;
            }

            if (!_imagesByKey.TryGetValue(key, out cached))
            {
                cached = SKImage.FromEncodedData(data);
                _imagesByKey[key] = cached;
            }

            _imagesByPictureVariant[variantKey] = cached;
            return cached;
        }

        private static int ToTargetPixels(float points)
        {
            var pixels = points * 96f / 72f;
            if (pixels < 1f)
            {
                pixels = 1f;
            }

            return (int)Math.Ceiling(pixels);
        }

        private static string BuildVariantKey(byte[] imageData, int targetWidthPx, int targetHeightPx)
        {
            return BuildKey(imageData)
                + ":"
                + targetWidthPx.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + "x"
                + targetHeightPx.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string BuildKey(byte[] imageData)
        {
            unchecked
            {
                uint hash = 2166136261;
                for (var i = 0; i < imageData.Length; i++)
                {
                    hash ^= imageData[i];
                    hash *= 16777619;
                }

                return imageData.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    + ":"
                    + hash.ToString("X8", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        public void Dispose()
        {
            foreach (var image in _imagesByKey.Values)
            {
                if (image != null)
                {
                    image.Dispose();
                }
            }

            foreach (var data in _dataByKey.Values)
            {
                if (data != null)
                {
                    data.Dispose();
                }
            }

            _imagesByPictureVariant.Clear();
            _imagesByKey.Clear();
            _dataByKey.Clear();
        }
    }
}
