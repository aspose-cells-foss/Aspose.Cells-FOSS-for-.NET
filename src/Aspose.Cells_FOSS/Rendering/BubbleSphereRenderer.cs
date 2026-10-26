using System;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal static class BubbleSphereRenderer
    {
        public static SKImage CreateTexture(SKColor color)
        {
            using (var surface = SKSurface.Create(new SKImageInfo(96, 96)))
            {
                surface.Canvas.Clear(SKColors.Transparent);
                Draw(surface.Canvas, new SKPoint(48f, 48f), 46f, color);
                return surface.Snapshot();
            }
        }

        public static void DrawImage(SKCanvas canvas, SKPoint center, float radius, SKImage texture)
        {
            var halfSide = radius * 48f / 46f;
            var bounds = new SKRect(center.X - halfSide, center.Y - halfSide,
                center.X + halfSide, center.Y + halfSide);
            using (var paint = new SKPaint { FilterQuality = SKFilterQuality.High,
                IsAntialias = true })
            {
                canvas.DrawImage(texture, bounds, paint);
            }
        }

        public static void Draw(SKCanvas canvas, SKPoint center, float radius, SKColor color)
        {
            if (radius <= 0f)
            {
                return;
            }

            float hue;
            float saturation;
            float value;
            color.ToHsv(out hue, out saturation, out value);
            var highlightHue = hue < 50f ? hue + 15f : hue;
            var highlightSaturation = hue < 50f ? saturation * 0.58f : saturation * 0.68f;
            var highlight = SKColor.FromHsv(highlightHue, highlightSaturation,
                Math.Min(100f, value + Math.Max(9f, (100f - value) * 0.6f)), color.Alpha);
            var lit = SKColor.FromHsv(hue, saturation * 0.98f,
                Math.Min(100f, value * 1.08f), color.Alpha);
            var shade = SKColor.FromHsv(hue, saturation * 0.9f, value * 0.48f, color.Alpha);
            var edge = SKColor.FromHsv(hue, saturation * 0.85f, value * 0.27f, color.Alpha);
            var lightCenter = new SKPoint(center.X - radius * 0.20f, center.Y - radius * 0.44f);
            var colors = new[] { highlight, lit, color, shade, edge };
            var stops = new[] { 0f, 0.20f, 0.33f, 0.65f, 1f };
            using (var shader = SKShader.CreateRadialGradient(lightCenter, radius * 1.65f,
                colors, stops, SKShaderTileMode.Clamp))
            using (var fill = new SKPaint { Style = SKPaintStyle.Fill, Shader = shader, IsAntialias = true })
            using (var reflectedLight = SKShader.CreateLinearGradient(
                center, new SKPoint(center.X, center.Y + radius),
                new[] { new SKColor(255, 255, 255, 0),
                    new SKColor(255, 255, 255, 55) },
                new[] { 0f, 1f }, SKShaderTileMode.Clamp))
            using (var reflection = new SKPaint { Style = SKPaintStyle.Fill,
                Shader = reflectedLight, IsAntialias = true })
            using (var specular = SKShader.CreateRadialGradient(lightCenter, radius * 0.42f,
                new[] { new SKColor(255, 255, 255, 48), SKColors.Transparent },
                SKShaderTileMode.Clamp))
            using (var glint = new SKPaint { Style = SKPaintStyle.Fill,
                Shader = specular, IsAntialias = true })
            {
                canvas.DrawCircle(center, radius, fill);
                canvas.DrawCircle(center, radius, reflection);
                canvas.DrawCircle(center, radius, glint);
            }
        }
    }
}
