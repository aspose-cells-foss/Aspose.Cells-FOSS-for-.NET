using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfTrueTypePrototypeUsage
    {
        private readonly SKTypeface _typeface;
        private readonly Dictionary<byte, float> _widths1000;
        private readonly Dictionary<byte, ushort> _codeToGlyphId;
        private readonly Dictionary<ushort, int> _glyphIdToUnicode;
        private readonly HashSet<byte> _codes;
        private readonly SKFont _font;
        private readonly SKPaint _paint;
        private readonly byte[] _fontBytes;
        private readonly TrueTypeSubsetFontProgram _subsetProgram;
        private SKRect _fontBounds;
        private bool _hasBounds;

        public PdfTrueTypePrototypeUsage(SKTypeface typeface)
        {
            if (typeface == null)
            {
                throw new ArgumentNullException("typeface");
            }

            _typeface = typeface;
            _widths1000 = new Dictionary<byte, float>();
            _codeToGlyphId = new Dictionary<byte, ushort>();
            _glyphIdToUnicode = new Dictionary<ushort, int>();
            _codes = new HashSet<byte>();
            _font = new SKFont(typeface, 1000f);
            _font.Subpixel = true;
            _paint = new SKPaint(_font);
            _paint.Style = SKPaintStyle.Fill;
            _paint.IsAntialias = true;

            _fontBytes = ReadFontBytes(typeface);
            _subsetProgram = new TrueTypeSubsetFontProgram(_fontBytes);
        }

        public SKTypeface Typeface
        {
            get { return _typeface; }
        }

        public string ResourceName { get; set; }

        public int FontObjectId { get; set; }

        public int FontDescriptorObjectId { get; set; }

        public int FontFileObjectId { get; set; }

        public int ToUnicodeObjectId { get; set; }

        public string SubsetFontName { get; set; }

        public byte[] FontBytes
        {
            get { return _fontBytes; }
        }

        public TrueTypeSubsetFontProgram SubsetProgram
        {
            get { return _subsetProgram; }
        }

        public string BaseFontName
        {
            get
            {
                var family = _typeface.FamilyName;
                if (string.IsNullOrEmpty(family))
                {
                    family = "SkiaFont";
                }

                var builder = new System.Text.StringBuilder();
                for (var i = 0; i < family.Length; i++)
                {
                    var ch = family[i];
                    if (char.IsLetterOrDigit(ch))
                    {
                        builder.Append(ch);
                    }
                }

                if (builder.Length == 0)
                {
                    builder.Append("SkiaFont");
                }

                return builder.ToString();
            }
        }

        public byte FirstChar
        {
            get
            {
                var first = (byte)255;
                foreach (var code in _codes)
                {
                    if (code < first)
                    {
                        first = code;
                    }
                }

                if (_codes.Count == 0)
                {
                    return 32;
                }

                return first;
            }
        }

        public byte LastChar
        {
            get
            {
                var last = (byte)32;
                foreach (var code in _codes)
                {
                    if (code > last)
                    {
                        last = code;
                    }
                }

                return last;
            }
        }

        public SKRect FontBounds
        {
            get { return _fontBounds; }
        }

        public bool HasBounds
        {
            get { return _hasBounds; }
        }

        public void AddText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                if (ch > 255)
                {
                    continue;
                }

                var code = (byte)ch;
                if (_codes.Add(code))
                {
                    var token = new string(ch, 1);
                    var glyphId = ResolveGlyphId(ch);
                    _codeToGlyphId[code] = glyphId;
                    if (!_glyphIdToUnicode.ContainsKey(glyphId))
                    {
                        _glyphIdToUnicode[glyphId] = ch;
                    }

                    _widths1000[code] = _paint.MeasureText(token);
                    using (var path = _paint.GetTextPath(token, 0f, 0f))
                    {
                        if (path != null && !path.IsEmpty)
                        {
                            var bounds = path.Bounds;
                            if (!_hasBounds)
                            {
                                _fontBounds = bounds;
                                _hasBounds = true;
                            }
                            else
                            {
                                _fontBounds = SKRect.Union(_fontBounds, bounds);
                            }
                        }
                    }
                }
            }
        }

        public byte[] EncodeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new byte[0];
            }

            var bytes = new List<byte>();
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                if (ch > 255)
                {
                    continue;
                }

                bytes.Add((byte)ch);
            }

            return bytes.ToArray();
        }

        public float WidthForCode(byte code)
        {
            float width;
            if (_widths1000.TryGetValue(code, out width))
            {
                return width;
            }

            return 0f;
        }

        public IList<byte> OrderedCodes()
        {
            var ordered = new List<byte>(_codes);
            ordered.Sort();
            return ordered;
        }

        public IList<ushort> OrderedGlyphIds()
        {
            var glyphIds = new List<ushort>(_codeToGlyphId.Values);
            glyphIds.Sort();
            return glyphIds;
        }

        public IDictionary<byte, ushort> CodeToGlyphIdMap()
        {
            return _codeToGlyphId;
        }

        public IDictionary<ushort, int> GlyphIdToUnicodeMap()
        {
            return _glyphIdToUnicode;
        }

        public byte[] BuildSubsetFontBytes()
        {
            return _subsetProgram.BuildSubset(OrderedGlyphIds(), _codeToGlyphId, _glyphIdToUnicode, SubsetFontName);
        }

        private ushort ResolveGlyphId(char ch)
        {
            return _font.GetGlyph((int)ch);
        }

        private static byte[] ReadFontBytes(SKTypeface typeface)
        {
            using (var stream = typeface.OpenStream())
            {
                if (stream == null || !stream.HasLength)
                {
                    throw new InvalidOperationException("Could not open a readable font stream for " + typeface.FamilyName + ".");
                }

                var length = stream.Length;
                if (length <= 0)
                {
                    throw new InvalidOperationException("Font stream length was empty for " + typeface.FamilyName + ".");
                }

                var bytes = new byte[length];
                var offset = 0;
                var chunk = new byte[8192];
                while (offset < length)
                {
                    var remaining = length - offset;
                    var readSize = remaining < chunk.Length ? remaining : chunk.Length;
                    var read = stream.Read(chunk, readSize);
                    if (read <= 0)
                    {
                        break;
                    }

                    Buffer.BlockCopy(chunk, 0, bytes, offset, read);
                    offset += read;
                }

                if (offset != length)
                {
                    throw new InvalidOperationException("Could not read the complete font stream for " + typeface.FamilyName + ".");
                }

                return bytes;
            }
        }

        public void Dispose()
        {
            if (_paint != null)
            {
                _paint.Dispose();
            }

            if (_font != null)
            {
                _font.Dispose();
            }
        }
    }
}
