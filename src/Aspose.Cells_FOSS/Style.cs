using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a mutable cell style facade that can be applied to one or more cells.
    /// </summary>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// var cell = workbook.Worksheets[0].Cells["C3"];
    ///
    /// var style = cell.GetStyle();
    /// style.Font.Bold = true;
    /// style.NumberFormat = "$#,##0.00";
    /// style.HorizontalAlignment = HorizontalAlignmentType.Center;
    /// cell.SetStyle(style);
    /// </code>
    /// </example>
    public class Style
    {
        private int _indentLevel;
        private int _textRotation;
        private int _readingOrder;

        /// <summary>
        /// Initializes a new style with default font and border objects.
        /// </summary>
        public Style()
        {
            Font = new Font();
            Borders = new Borders();
        }

        /// <summary>
        /// Gets or sets the font settings.
        /// </summary>
        public Font Font { get; set; }

        /// <summary>
        /// Gets or sets border settings.
        /// </summary>
        public Borders Borders { get; set; }

        /// <summary>
        /// Gets or sets the fill pattern.
        /// </summary>
        public FillPattern Pattern { get; set; }

        /// <summary>
        /// Gets or sets the fill foreground color.
        /// </summary>
        public Color ForegroundColor { get; set; } = Color.Empty;

        /// <summary>
        /// Gets or sets the fill background color.
        /// </summary>
        public Color BackgroundColor { get; set; } = Color.Empty;

        /// <summary>
        /// Gets or sets the numeric format identifier.
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// Gets or sets the custom number format code.
        /// </summary>
        public string Custom { get; set; }

        /// <summary>
        /// Gets or sets the resolved number format string.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        public HorizontalAlignmentType HorizontalAlignment { get; set; }

        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        public VerticalAlignmentType VerticalAlignment { get; set; } = VerticalAlignmentType.Bottom;

        /// <summary>
        /// Gets or sets whether text wraps within the cell.
        /// </summary>
        public bool WrapText { get; set; }

        /// <summary>
        /// Gets or sets the indentation level.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the text rotation.
        /// </summary>
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

        /// <summary>
        /// Gets or sets whether the cell content shrinks to fit.
        /// </summary>
        public bool ShrinkToFit { get; set; }

        /// <summary>
        /// Gets or sets the reading order.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the relative indent.
        /// </summary>
        public int RelativeIndent { get; set; }

        /// <summary>
        /// Gets or sets whether the cell is locked when worksheet protection is enabled.
        /// </summary>
        public bool IsLocked { get; set; } = true;

        /// <summary>
        /// Gets or sets whether the cell formula is hidden when worksheet protection is enabled.
        /// </summary>
        public bool IsHidden { get; set; }

        /// <summary>
        /// Creates a copy of the current style.
        /// </summary>
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

        internal static Style FromCore(StyleValue value)
        {
            if (value == null)
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
}
