using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public class Style
{
    private int _indentLevel;
    private int _textRotation;
    private int _readingOrder;

    public Style()
    {
        Font = new Font();
        Borders = new Borders();
    }

    public Font Font { get; set; }
    public Borders Borders { get; set; }
    public FillPattern Pattern { get; set; }
    public Color ForegroundColor { get; set; } = Color.Empty;
    public Color BackgroundColor { get; set; } = Color.Empty;
    public int Number { get; set; }
    public string? Custom { get; set; }
    public string NumberFormat
    {
        get
        {
            return Aspose.Cells_FOSS.NumberFormat.ResolveFormatCode(Number, Custom);
        }
        set
        {
            var builtInId = Aspose.Cells_FOSS.NumberFormat.GetBuiltInFormatId(value);
            if (builtInId.HasValue)
            {
                Number = builtInId.Value;
                Custom = null;
                return;
            }

            Number = 0;
            Custom = string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }
    public HorizontalAlignmentType HorizontalAlignment { get; set; }
    public VerticalAlignmentType VerticalAlignment { get; set; } = VerticalAlignmentType.Bottom;
    public bool WrapText { get; set; }
    public int IndentLevel
    {
        get
        {
            return _indentLevel;
        }
        set
        {
            if (value < 0 || value > 250)
            {
                throw new CellsException("IndentLevel must be between 0 and 250.");
            }

            _indentLevel = value;
        }
    }
    public int TextRotation
    {
        get
        {
            return _textRotation;
        }
        set
        {
            if (value != 255 && (value < 0 || value > 180))
            {
                throw new CellsException("TextRotation must be between 0 and 180, or 255 for vertical text.");
            }

            _textRotation = value;
        }
    }
    public bool ShrinkToFit { get; set; }
    public int ReadingOrder
    {
        get
        {
            return _readingOrder;
        }
        set
        {
            if (value < 0 || value > 2)
            {
                throw new CellsException("ReadingOrder must be 0, 1, or 2.");
            }

            _readingOrder = value;
        }
    }
    public int RelativeIndent { get; set; }
    public bool IsLocked { get; set; } = true;
    public bool IsHidden { get; set; }

    public Style Clone()
    {
        return new Style
        {
            Font = Font.Clone(),
            Borders = Borders.Clone(),
            Pattern = Pattern,
            ForegroundColor = ForegroundColor,
            BackgroundColor = BackgroundColor,
            Number = Number,
            Custom = Custom,
            HorizontalAlignment = HorizontalAlignment,
            VerticalAlignment = VerticalAlignment,
            WrapText = WrapText,
            IndentLevel = IndentLevel,
            TextRotation = TextRotation,
            ShrinkToFit = ShrinkToFit,
            ReadingOrder = ReadingOrder,
            RelativeIndent = RelativeIndent,
            IsLocked = IsLocked,
            IsHidden = IsHidden,
        };
    }

    internal StyleValue ToCore()
    {
        return new StyleValue
        {
            Font = new FontValue
            {
                Name = Font.Name,
                Size = Font.Size,
                Bold = Font.Bold,
                Italic = Font.Italic,
                Underline = Font.Underline,
                StrikeThrough = Font.StrikeThrough,
                Color = Font.Color.ToCore(),
            },
            Pattern = (FillPatternKind)Pattern,
            ForegroundColor = ForegroundColor.ToCore(),
            BackgroundColor = BackgroundColor.ToCore(),
            Borders = new BordersValue
            {
                Left = ToCoreBorder(Borders.Left),
                Right = ToCoreBorder(Borders.Right),
                Top = ToCoreBorder(Borders.Top),
                Bottom = ToCoreBorder(Borders.Bottom),
                Diagonal = ToCoreBorder(Borders.Diagonal),
                DiagonalUp = Borders.DiagonalUp,
                DiagonalDown = Borders.DiagonalDown,
            },
            Alignment = new AlignmentValue
            {
                Horizontal = (Aspose.Cells_FOSS.Core.HorizontalAlignment)HorizontalAlignment,
                Vertical = (Aspose.Cells_FOSS.Core.VerticalAlignment)VerticalAlignment,
                WrapText = WrapText,
                IndentLevel = IndentLevel,
                TextRotation = TextRotation,
                ShrinkToFit = ShrinkToFit,
                ReadingOrder = ReadingOrder,
                RelativeIndent = RelativeIndent,
            },
            Protection = new ProtectionValue
            {
                IsLocked = IsLocked,
                IsHidden = IsHidden,
            },
            NumberFormat = new NumberFormatValue
            {
                Number = Number,
                Custom = Custom,
            },
        };
    }

    internal static Style FromCore(StyleValue? value)
    {
        if (value is null)
        {
            return new Style();
        }

        return new Style
        {
            Font = new Font
            {
                Name = value.Font.Name,
                Size = value.Font.Size,
                Bold = value.Font.Bold,
                Italic = value.Font.Italic,
                Underline = value.Font.Underline,
                StrikeThrough = value.Font.StrikeThrough,
                Color = Color.FromCore(value.Font.Color),
            },
            Pattern = (FillPattern)value.Pattern,
            ForegroundColor = Color.FromCore(value.ForegroundColor),
            BackgroundColor = Color.FromCore(value.BackgroundColor),
            Borders = new Borders
            {
                Left = FromCoreBorder(value.Borders.Left),
                Right = FromCoreBorder(value.Borders.Right),
                Top = FromCoreBorder(value.Borders.Top),
                Bottom = FromCoreBorder(value.Borders.Bottom),
                Diagonal = FromCoreBorder(value.Borders.Diagonal),
                DiagonalUp = value.Borders.DiagonalUp,
                DiagonalDown = value.Borders.DiagonalDown,
            },
            HorizontalAlignment = (HorizontalAlignmentType)value.Alignment.Horizontal,
            VerticalAlignment = (VerticalAlignmentType)value.Alignment.Vertical,
            WrapText = value.Alignment.WrapText,
            IndentLevel = value.Alignment.IndentLevel,
            TextRotation = value.Alignment.TextRotation,
            ShrinkToFit = value.Alignment.ShrinkToFit,
            ReadingOrder = value.Alignment.ReadingOrder,
            RelativeIndent = value.Alignment.RelativeIndent,
            IsLocked = value.Protection.IsLocked,
            IsHidden = value.Protection.IsHidden,
            Number = value.NumberFormat.Number,
            Custom = value.NumberFormat.Custom,
        };
    }

    private static BorderSideValue ToCoreBorder(Border border)
    {
        return new BorderSideValue
        {
            Style = (BorderStyle)border.LineStyle,
            Color = border.Color.ToCore(),
        };
    }

    private static Border FromCoreBorder(BorderSideValue value)
    {
        return new Border
        {
            LineStyle = (BorderStyleType)value.Style,
            Color = Color.FromCore(value.Color),
        };
    }
}
