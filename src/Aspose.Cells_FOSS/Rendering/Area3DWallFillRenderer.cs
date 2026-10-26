using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class Area3DWallFillRenderer
    {
        public static void Draw(
            SKCanvas canvas,
            SKPoint frontBottomLeft,
            SKPoint frontTopLeft,
            SKPoint backBottomLeft,
            SKPoint backTopLeft,
            SKPoint backBottomRight,
            SKPoint backTopRight,
            byte[] sideWallImageData,
            byte[] backWallImageData,
            SKColor fallbackColor)
        {
            using (var sideWall = new SKPath())
            {
                sideWall.MoveTo(frontTopLeft);
                sideWall.LineTo(backTopLeft);
                sideWall.LineTo(backBottomLeft);
                sideWall.LineTo(frontBottomLeft);
                sideWall.Close();
                DrawWall(canvas, sideWall, sideWallImageData, fallbackColor, 0.72f);
            }

            using (var backWall = new SKPath())
            {
                backWall.MoveTo(backTopLeft);
                backWall.LineTo(backTopRight);
                backWall.LineTo(backBottomRight);
                backWall.LineTo(backBottomLeft);
                backWall.Close();
                DrawWall(canvas, backWall, backWallImageData, fallbackColor, 1f);
            }
        }

        public static void DrawFloor(
            SKCanvas canvas,
            SKPoint backLeft,
            SKPoint backRight,
            SKPoint frontRight,
            SKPoint frontLeft,
            byte[] imageData,
            SKColor fallbackColor)
        {
            using (var floor = new SKPath())
            {
                floor.MoveTo(backLeft);
                floor.LineTo(backRight);
                floor.LineTo(frontRight);
                floor.LineTo(frontLeft);
                floor.Close();
                DrawWall(canvas, floor, imageData, fallbackColor, 0.88f);
            }
        }

        private static void DrawWall(SKCanvas canvas, SKPath wall, byte[] imageData, SKColor fallbackColor, float brightness)
        {
            using (var fallback = new SKPaint
            {
                Style = SKPaintStyle.Fill,
                Color = Shade(fallbackColor, brightness),
                IsAntialias = true
            })
            {
                canvas.DrawPath(wall, fallback);
            }

            if (imageData == null || imageData.Length == 0)
            {
                return;
            }

            using (var data = SKData.CreateCopy(imageData))
            using (var image = SKImage.FromEncodedData(data))
            {
                DrawTiles(canvas, wall, image);
            }

            if (brightness < 0.999f)
            {
                var alpha = (byte)Math.Round((1f - brightness) * 255f);
                using (var shadow = new SKPaint
                {
                    Style = SKPaintStyle.Fill,
                    Color = new SKColor(0, 0, 0, alpha),
                    IsAntialias = true
                })
                {
                    canvas.DrawPath(wall, shadow);
                }
            }
        }

        private static void DrawTiles(SKCanvas canvas, SKPath wall, SKImage image)
        {
            if (image == null || image.Width <= 0 || image.Height <= 0)
            {
                return;
            }

            var bounds = wall.Bounds;
            canvas.Save();
            canvas.ClipPath(wall, SKClipOperation.Intersect, true);
            using (var paint = new SKPaint { IsAntialias = true, FilterQuality = SKFilterQuality.High })
            {
                var tileWidth = image.Width * 72f / 96f;
                var tileHeight = image.Height * 72f / 96f;
                var firstX = bounds.Left - PositiveRemainder(bounds.Left, tileWidth);
                var firstY = bounds.Top - PositiveRemainder(bounds.Top, tileHeight);
                for (var y = firstY; y < bounds.Bottom; y += tileHeight)
                {
                    for (var x = firstX; x < bounds.Right; x += tileWidth)
                    {
                        var destination = new SKRect(x, y, x + tileWidth, y + tileHeight);
                        canvas.DrawImage(image, destination, paint);
                    }
                }
            }
            canvas.Restore();
        }

        private static float PositiveRemainder(float value, float divisor)
        {
            var remainder = value % divisor;
            return remainder < 0f ? remainder + divisor : remainder;
        }

        private static SKColor Shade(SKColor color, float brightness)
        {
            return new SKColor(
                (byte)Math.Round(color.Red * brightness),
                (byte)Math.Round(color.Green * brightness),
                (byte)Math.Round(color.Blue * brightness),
                color.Alpha);
        }
    }
}
