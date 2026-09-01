using System;
using System.Collections.Generic;
using System.Globalization;
using SkiaSharp;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class PdfType3FontUsage : IDisposable
    {
        private readonly SKTypeface _typeface;
        private readonly bool _usePositiveYFontMatrix;
        private readonly Dictionary<string, float> _tokenWidths = new Dictionary<string, float>(StringComparer.Ordinal);
        private readonly List<PdfType3FontSubset> _subsets = new List<PdfType3FontSubset>();
        private readonly SKPaint _paint;

        public PdfType3FontUsage(SKTypeface typeface, bool usePositiveYFontMatrix)
        {
            _typeface = typeface;
            _usePositiveYFontMatrix = usePositiveYFontMatrix;
            using (var font = new SKFont(typeface, 1000f))
            {
                font.Subpixel = true;
                _paint = new SKPaint(font);
                _paint.Style = SKPaintStyle.Fill;
                _paint.IsAntialias = true;
            }
        }

        public SKTypeface Typeface
        {
            get { return _typeface; }
        }

        public bool UsePositiveYFontMatrix
        {
            get { return _usePositiveYFontMatrix; }
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

        public IList<PdfType3FontSubset> Subsets
        {
            get { return _subsets; }
        }

        public void AddText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var index = 0;
            while (index < text.Length)
            {
                var token = ReadToken(text, ref index);
                if (!_tokenWidths.ContainsKey(token))
                {
                    _tokenWidths[token] = _paint.MeasureText(token);
                }
            }
        }

        public void FinalizeSubsets(ref int resourceIndex)
        {
            if (_subsets.Count > 0)
            {
                return;
            }

            var orderedTokens = new List<string>(_tokenWidths.Keys);
            orderedTokens.Sort(StringComparer.Ordinal);

            PdfType3FontSubset current = null;
            for (var i = 0; i < orderedTokens.Count; i++)
            {
                var token = orderedTokens[i];
                if (current == null || current.Tokens.Count >= 255)
                {
                    current = new PdfType3FontSubset();
                    current.ResourceName = "T3F" + resourceIndex.ToString(CultureInfo.InvariantCulture);
                    resourceIndex++;
                    _subsets.Add(current);
                }

                var code = (byte)(current.Tokens.Count + 1);
                current.Tokens.Add(token);
                current.TokenCodes[token] = code;
            }
        }

        public bool TryResolveToken(string token, out PdfType3FontSubset subset, out byte code, out float width1000)
        {
            width1000 = 0f;
            subset = null;
            code = 0;

            if (!_tokenWidths.TryGetValue(token, out width1000))
            {
                return false;
            }

            for (var i = 0; i < _subsets.Count; i++)
            {
                byte resolvedCode;
                if (_subsets[i].TokenCodes.TryGetValue(token, out resolvedCode))
                {
                    subset = _subsets[i];
                    code = resolvedCode;
                    return true;
                }
            }

            return false;
        }

        public float GetTokenWidth(string token)
        {
            float width;
            if (_tokenWidths.TryGetValue(token, out width))
            {
                return width;
            }

            return 0f;
        }

        public SKPath BuildTokenPath(string token)
        {
            return _paint.GetTextPath(token, 0f, 0f);
        }

        private static string ReadToken(string text, ref int index)
        {
            var ch = text[index];
            if (char.IsHighSurrogate(ch) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]))
            {
                var token = text.Substring(index, 2);
                index += 2;
                return token;
            }

            index++;
            return text.Substring(index - 1, 1);
        }

        public void Dispose()
        {
            if (_paint != null)
            {
                _paint.Dispose();
            }
        }
    }
}
