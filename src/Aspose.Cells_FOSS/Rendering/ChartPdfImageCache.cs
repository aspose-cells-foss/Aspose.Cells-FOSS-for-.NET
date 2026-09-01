using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>
    /// Renders charts to cached bitmap snapshots for PDF embedding. Line-heavy charts need a higher
    /// raster fidelity than photos, otherwise axis labels, gridlines, and thin series strokes pick
    /// up JPEG softness.
    /// </summary>
    internal sealed class ChartPdfImageCache : IDisposable
    {
        private readonly Dictionary<string, SKData> _dataByKey = new Dictionary<string, SKData>(StringComparer.Ordinal);
        private readonly Dictionary<string, SKImage> _imagesByKey = new Dictionary<string, SKImage>(StringComparer.Ordinal);

        public SKImage GetImage(ChartModel chart, ParsedChart parsed, SKRect rect, SKRect visualRect, ChartRenderer renderer)
        {
            if (chart == null || parsed == null || renderer == null || rect.Width <= 0f || rect.Height <= 0f)
            {
                return null;
            }

            var targetDpi = ResolveTargetDpi(parsed, rect);
            var encodeLosslessly = ShouldEncodeLosslessly(parsed);
            var jpegQuality = ResolveJpegQuality(parsed, rect, targetDpi);
            var widthPx = ToTargetPixels(rect.Width, targetDpi);
            var heightPx = ToTargetPixels(rect.Height, targetDpi);
            var cropBounds = ResolveCropBounds(rect, visualRect, widthPx, heightPx);
            var key = BuildKey(chart, widthPx, heightPx, jpegQuality, encodeLosslessly, cropBounds);

            SKImage cached;
            if (_imagesByKey.TryGetValue(key, out cached))
            {
                return cached;
            }

            using (var bitmap = new SKBitmap(widthPx, heightPx, true))
            using (var canvas = new SKCanvas(bitmap))
            {
                canvas.Clear(SKColors.White);
                var scaleX = widthPx / rect.Width;
                var scaleY = heightPx / rect.Height;
                canvas.Scale(scaleX, scaleY);
                renderer.Draw(canvas, new SKRect(0f, 0f, rect.Width, rect.Height), parsed);

                using (var image = SKImage.FromBitmap(bitmap))
                {
                    if (image == null)
                    {
                        return null;
                    }

                    SKData encoded;
                    using (var croppedBitmap = CropBitmap(bitmap, cropBounds))
                    using (var croppedImage = SKImage.FromBitmap(croppedBitmap))
                    {
                        if (croppedImage == null)
                        {
                            return null;
                        }

                        if (encodeLosslessly)
                        {
                            encoded = croppedImage.Encode(SKEncodedImageFormat.Png, 100);
                        }
                        else
                        {
                            encoded = croppedImage.Encode(SKEncodedImageFormat.Jpeg, jpegQuality);
                        }
                    }

                    if (encoded == null || encoded.Size <= 0)
                    {
                        return null;
                    }

                    cached = SKImage.FromEncodedData(encoded);
                    if (cached == null)
                    {
                        encoded.Dispose();
                        return null;
                    }

                    _dataByKey[key] = encoded;
                    _imagesByKey[key] = cached;
                    return cached;
                }
            }
        }

        private static int ToTargetPixels(float points, float targetDpi)
        {
            var pixels = points * targetDpi / 72f;
            if (pixels < 1f)
            {
                pixels = 1f;
            }

            return (int)Math.Ceiling(pixels);
        }

        private static float ResolveTargetDpi(ParsedChart parsed, SKRect rect)
        {
            var area = rect.Width * rect.Height;
            if (parsed != null && (parsed.Kind == ChartKind.Line || parsed.Kind == ChartKind.Area))
            {
                var pointCount = MaxPointCount(parsed);
                if (pointCount <= 12 && area >= 50000f)
                {
                    return 192f;
                }

                return 216f;
            }

            if (parsed != null && (parsed.Kind == ChartKind.Column || parsed.Kind == ChartKind.Bar))
            {
                if (area >= 50000f)
                {
                    return 180f;
                }

                return 192f;
            }

            if (area >= 70000f)
            {
                return 144f;
            }

            return 168f;
        }

        private static int ResolveJpegQuality(ParsedChart parsed, SKRect rect, float targetDpi)
        {
            if (parsed != null && parsed.Kind == ChartKind.Line)
            {
                if (targetDpi <= 120f)
                {
                    return 58;
                }

                return 66;
            }

            if (parsed != null && parsed.Kind == ChartKind.Area)
            {
                return targetDpi <= 150f ? 68 : 74;
            }

            return targetDpi <= 144f ? 70 : 76;
        }

        private static bool ShouldEncodeLosslessly(ParsedChart parsed)
        {
            if (parsed == null)
            {
                return false;
            }

            return parsed.Kind == ChartKind.Line
                || parsed.Kind == ChartKind.Area
                || parsed.Kind == ChartKind.Column
                || parsed.Kind == ChartKind.Bar;
        }

        private static int MaxPointCount(ParsedChart chart)
        {
            var count = chart != null ? chart.Categories.Count : 0;
            if (chart == null)
            {
                return count;
            }

            for (var i = 0; i < chart.Series.Count; i++)
            {
                if (chart.Series[i] != null && chart.Series[i].Values.Count > count)
                {
                    count = chart.Series[i].Values.Count;
                }
            }

            return count;
        }

        private static SKRectI ResolveCropBounds(SKRect rect, SKRect visualRect, int widthPx, int heightPx)
        {
            if (visualRect.Width <= 0f || visualRect.Height <= 0f)
            {
                return new SKRectI(0, 0, widthPx, heightPx);
            }

            var scaleX = widthPx / rect.Width;
            var scaleY = heightPx / rect.Height;
            var left = (int)Math.Floor((visualRect.Left - rect.Left) * scaleX);
            var top = (int)Math.Floor((visualRect.Top - rect.Top) * scaleY);
            var right = (int)Math.Ceiling((visualRect.Right - rect.Left) * scaleX);
            var bottom = (int)Math.Ceiling((visualRect.Bottom - rect.Top) * scaleY);

            if (left < 0)
            {
                left = 0;
            }

            if (top < 0)
            {
                top = 0;
            }

            if (right > widthPx)
            {
                right = widthPx;
            }

            if (bottom > heightPx)
            {
                bottom = heightPx;
            }

            if (right <= left)
            {
                left = 0;
                right = widthPx;
            }

            if (bottom <= top)
            {
                top = 0;
                bottom = heightPx;
            }

            return new SKRectI(left, top, right, bottom);
        }

        private static SKBitmap CropBitmap(SKBitmap bitmap, SKRectI cropBounds)
        {
            var width = cropBounds.Right - cropBounds.Left;
            var height = cropBounds.Bottom - cropBounds.Top;
            if (width <= 0 || height <= 0)
            {
                return bitmap.Copy();
            }

            var result = new SKBitmap(width, height, true);
            using (var canvas = new SKCanvas(result))
            {
                canvas.Clear(SKColors.White);
                var source = new SKRect(cropBounds.Left, cropBounds.Top, cropBounds.Right, cropBounds.Bottom);
                var destination = new SKRect(0f, 0f, width, height);
                canvas.DrawBitmap(bitmap, source, destination);
            }

            return result;
        }

        private static string BuildKey(ChartModel chart, int widthPx, int heightPx, int jpegQuality, bool encodeLosslessly, SKRectI cropBounds)
        {
            unchecked
            {
                var hash = 17;
                hash = hash * 31 + (chart.RawChartXml != null ? chart.RawChartXml.GetHashCode() : 0);
                hash = hash * 31 + chart.UpperLeftRow;
                hash = hash * 31 + chart.UpperLeftColumn;
                hash = hash * 31 + chart.LowerRightRow;
                hash = hash * 31 + chart.LowerRightColumn;
                hash = hash * 31 + widthPx;
                hash = hash * 31 + heightPx;
                hash = hash * 31 + jpegQuality;
                hash = hash * 31 + (encodeLosslessly ? 1 : 0);
                hash = hash * 31 + cropBounds.Left;
                hash = hash * 31 + cropBounds.Top;
                hash = hash * 31 + cropBounds.Right;
                hash = hash * 31 + cropBounds.Bottom;
                return hash.ToString(System.Globalization.CultureInfo.InvariantCulture);
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

            _imagesByKey.Clear();
            _dataByKey.Clear();
        }
    }
}
