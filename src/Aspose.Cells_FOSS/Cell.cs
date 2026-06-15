using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a single worksheet cell and exposes value, formula, and style operations.
    /// </summary>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// var cell = workbook.Worksheets[0].Cells["B2"];
    ///
    /// cell.PutValue(12.5);
    /// cell.Formula = "=SUM(B2:B10)";
    ///
    /// var style = cell.GetStyle();
    /// style.NumberFormat = "$#,##0.00";
    /// cell.SetStyle(style);
    /// </code>
    /// </example>
    public class Cell
    {
        private readonly Worksheet _worksheet;
        private readonly CellAddress _address;

        internal Cell(Worksheet worksheet, CellAddress address)
        {
            _worksheet = worksheet;
            _address = address;
        }

        /// <summary>
        /// Gets or sets the logical cell value.
        /// </summary>
        /// <remarks>
        /// Supported assignments in v0.1 include strings, booleans, numbers, <see cref="DateTime"/>, and <see langword="null"/>.
        /// </remarks>
        public object Value
        {
            get
            {
                var record = TryGetRecord();
                return record?.Value;
            }
            set
            {
                AssignValue(value);
            }
        }

        /// <summary>
        /// Gets a stable string representation of the cell value without applying style-based display formatting.
        /// </summary>
        public string StringValue
        {
            get
            {
                var record = TryGetRecord();
                return DisplayTextFormatter.FormatStringValue(record?.Value);
            }
        }

        /// <summary>
        /// Gets the display text generated from the cell value, style, and workbook culture.
        /// </summary>
        public string DisplayStringValue
        {
            get
            {
                var record = TryGetRecord();
                // Display formatting is style-aware while StringValue stays a stable raw
                // representation of the logical cell payload.
                var style = record == null ? _worksheet.Workbook.Model.DefaultStyle : record.Style;
                return DisplayTextFormatter.FormatDisplayValue(record?.Value, style, _worksheet.Workbook.Model.Settings.DisplayCulture);
            }
        }

        /// <summary>
        /// Gets or sets the cell formula.
        /// </summary>
        /// <remarks>
        /// Formulas are stored and round-tripped in v0.1, but they are not recalculated automatically.
        /// </remarks>
        public string Formula
        {
            get
            {
                var record = TryGetRecord();
                if (record == null || string.IsNullOrEmpty(record.Formula))
                {
                    return string.Empty;
                }

                return "=" + record.Formula;
            }
            set
            {
                var record = GetOrCreateRecord();
                // Store formulas without a leading '=' so XML persistence and comparisons
                // have one normalized internal representation.
                record.Formula = NormalizeFormula(value);
                if (string.IsNullOrEmpty(record.Formula))
                {
                    if (record.Value == null)
                    {
                        record.Kind = CellValueKind.Blank;
                    }

                    return;
                }

                record.Kind = CellValueKind.Formula;
            }
        }

        /// <summary>
        /// Gets the current logical cell value type.
        /// </summary>
        public CellValueType Type
        {
            get
            {
                var record = TryGetRecord();
                if (record == null)
                {
                    return CellValueType.IsNull;
                }

                if (!string.IsNullOrEmpty(record.Formula))
                {
                    return CellValueType.IsUnknown;
                }

                switch (record.Kind)
                {
                    case CellValueKind.String:
                        return CellValueType.IsString;
                    case CellValueKind.Number:
                        return CellValueType.IsNumeric;
                    case CellValueKind.Boolean:
                        return CellValueType.IsBool;
                    case CellValueKind.DateTime:
                        return CellValueType.IsDateTime;
                    case CellValueKind.Error:
                        return CellValueType.IsError;
                    case CellValueKind.Formula:
                        return CellValueType.IsUnknown;
                    default:
                        return CellValueType.IsNull;
                }
            }
        }

        /// <summary>
        /// Sets the cell value to a string.
        /// </summary>
        public void PutValue(string value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            SetScalar(value, CellValueKind.String);
        }

        /// <summary>
        /// Sets the cell value to an integer.
        /// </summary>
        public void PutValue(int value)
        {
            SetScalar(value, CellValueKind.Number);
        }

        /// <summary>
        /// Sets the cell value to a floating-point number.
        /// </summary>
        public void PutValue(double value)
        {
            SetScalar(value, CellValueKind.Number);
        }

        /// <summary>
        /// Sets the cell value to a decimal number.
        /// </summary>
        public void PutValue(decimal value)
        {
            SetScalar(value, CellValueKind.Number);
        }

        /// <summary>
        /// Sets the cell value to a boolean.
        /// </summary>
        public void PutValue(bool value)
        {
            SetScalar(value, CellValueKind.Boolean);
        }

        /// <summary>
        /// Sets the cell value to a <see cref="DateTime"/>.
        /// </summary>
        /// <remarks>
        /// Date serialization honors <see cref="WorkbookSettings.Date1904"/> when the workbook is saved.
        /// </remarks>
        public void PutValue(DateTime value)
        {
            SetScalar(value, CellValueKind.DateTime);
        }

        /// <summary>
        /// Puts a string value into the cell and converts the value when appropriate.
        /// </summary>
        public void PutValue(string value, bool isConverted)
        {
            PutValue(value, isConverted, true);
        }

        /// <summary>
        /// Puts a string value into the cell and converts the value when appropriate.
        /// </summary>
        public void PutValue(string value, bool isConverted, bool setStyle)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (setStyle)
            {
                // Keep API-compatible signature. Number-format side effects are not
                // applied in the current v0.1 implementation.
            }

            if (!isConverted)
            {
                PutValue(value);
                return;
            }

            int intValue;
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out intValue))
            {
                PutValue(intValue);
                return;
            }

            double doubleValue;
            if (double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out doubleValue))
            {
                PutValue(doubleValue);
                return;
            }

            bool boolValue;
            if (bool.TryParse(value, out boolValue))
            {
                PutValue(boolValue);
                return;
            }

            DateTime dateTimeValue;
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTimeValue))
            {
                PutValue(dateTimeValue);
                return;
            }

            PutValue(value);
        }

        /// <summary>
        /// Puts an object value into the cell.
        /// </summary>
        public void PutValue(object value)
        {
            AssignValue(value);
        }

        /// <summary>
        /// Gets a detached copy of the cell style.
        /// </summary>
        public Style GetStyle()
        {
            var record = TryGetRecord();
            var style = record?.Style ?? _worksheet.Workbook.Model.DefaultStyle;
            return Style.FromCore(style.Clone());
        }

        /// <summary>
        /// Gets a detached copy of the cell style.
        /// </summary>
        /// <param name="checkBorders">Whether to consider surrounding border effects.</param>
        public Style GetStyle(bool checkBorders)
        {
            if (checkBorders)
            {
                // Kept for Aspose-compatible overload shape. Merge-aware resolution
                // is not differentiated in the current implementation.
            }

            return GetStyle();
        }

        /// <summary>
        /// Replaces the cell style with the supplied style object.
        /// </summary>
        public void SetStyle(Style style)
        {
            if (style == null) throw new ArgumentNullException(nameof(style));
            GetOrCreateRecord().Style = style.ToCore();
        }

        /// <summary>
        /// Applies the cell style with Aspose-compatible overload shape.
        /// </summary>
        /// <param name="style">The cell style.</param>
        /// <param name="explicitFlag">Reserved for explicit-only application semantics.</param>
        public void SetStyle(Style style, bool explicitFlag)
        {
            if (explicitFlag)
            {
                // Kept for Aspose-compatible overload shape. Explicit-only style
                // application is not differentiated in the current implementation.
            }

            SetStyle(style);
        }

        /// <summary>
        /// Applies the cell style based on style flags.
        /// </summary>
        /// <param name="style">The cell style.</param>
        /// <param name="flag">The style flag.</param>
        public void SetStyle(Style style, StyleFlag flag)
        {
            if (style == null) throw new ArgumentNullException(nameof(style));
            if (flag == null) throw new ArgumentNullException(nameof(flag));

            if (flag.All)
            {
                SetStyle(style);
                return;
            }

            var target = GetStyle();

            if (flag.NumberFormat)
            {
                target.Number = style.Number;
                target.Custom = style.Custom;
            }

            if (flag.CellShading)
            {
                target.Pattern = style.Pattern;
                target.ForegroundColor = style.ForegroundColor;
                target.BackgroundColor = style.BackgroundColor;
            }

            if (flag.Font)
            {
                target.Font.Name = style.Font.Name;
                target.Font.Size = style.Font.Size;
                target.Font.IsBold = style.Font.IsBold;
                target.Font.IsItalic = style.Font.IsItalic;
                target.Font.Underline = style.Font.Underline;
                target.Font.IsStrikeout = style.Font.IsStrikeout;
                target.Font.Color = style.Font.Color;
            }
            else
            {
                if (flag.FontName)
                {
                    target.Font.Name = style.Font.Name;
                }

                if (flag.FontSize)
                {
                    target.Font.Size = style.Font.Size;
                }

                if (flag.FontBold)
                {
                    target.Font.IsBold = style.Font.IsBold;
                }

                if (flag.FontItalic)
                {
                    target.Font.IsItalic = style.Font.IsItalic;
                }

                if (flag.FontUnderline)
                {
                    target.Font.Underline = style.Font.Underline;
                }

                if (flag.FontStrike)
                {
                    target.Font.IsStrikeout = style.Font.IsStrikeout;
                }

                if (flag.FontColor)
                {
                    target.Font.Color = style.Font.Color;
                }
            }

            if (flag.Borders)
            {
                target.Borders.Left = style.Borders.Left.Clone();
                target.Borders.Right = style.Borders.Right.Clone();
                target.Borders.Top = style.Borders.Top.Clone();
                target.Borders.Bottom = style.Borders.Bottom.Clone();
                target.Borders.Diagonal = style.Borders.Diagonal.Clone();
                target.Borders.DiagonalUp = style.Borders.DiagonalUp;
                target.Borders.DiagonalDown = style.Borders.DiagonalDown;
            }
            else
            {
                if (flag.LeftBorder)
                {
                    target.Borders.Left = style.Borders.Left.Clone();
                }

                if (flag.RightBorder)
                {
                    target.Borders.Right = style.Borders.Right.Clone();
                }

                if (flag.TopBorder)
                {
                    target.Borders.Top = style.Borders.Top.Clone();
                }

                if (flag.BottomBorder)
                {
                    target.Borders.Bottom = style.Borders.Bottom.Clone();
                }

                if (flag.DiagonalUpBorder)
                {
                    target.Borders.Diagonal = style.Borders.Diagonal.Clone();
                    target.Borders.DiagonalUp = style.Borders.DiagonalUp;
                }

                if (flag.DiagonalDownBorder)
                {
                    target.Borders.Diagonal = style.Borders.Diagonal.Clone();
                    target.Borders.DiagonalDown = style.Borders.DiagonalDown;
                }
            }

            if (flag.Alignments)
            {
                target.HorizontalAlignment = style.HorizontalAlignment;
                target.VerticalAlignment = style.VerticalAlignment;
                target.WrapText = style.WrapText;
                target.IndentLevel = style.IndentLevel;
                target.TextRotation = style.TextRotation;
                target.ReadingOrder = style.ReadingOrder;
                target.RelativeIndent = style.RelativeIndent;
                target.ShrinkToFit = style.ShrinkToFit;
            }
            else
            {
                if (flag.HorizontalAlignment)
                {
                    target.HorizontalAlignment = style.HorizontalAlignment;
                }

                if (flag.Indent)
                {
                    target.IndentLevel = style.IndentLevel;
                }

                if (flag.VerticalAlignment)
                {
                    target.VerticalAlignment = style.VerticalAlignment;
                }

                if (flag.WrapText)
                {
                    target.WrapText = style.WrapText;
                }

                if (flag.Rotation)
                {
                    target.TextRotation = style.TextRotation;
                }

                if (flag.TextDirection)
                {
                    target.ReadingOrder = style.ReadingOrder;
                }

                if (flag.ShrinkToFit)
                {
                    target.ShrinkToFit = style.ShrinkToFit;
                }
            }

            if (flag.Locked)
            {
                target.IsLocked = style.IsLocked;
            }

            if (flag.Hidden)
            {
                target.IsHidden = style.IsHidden;
            }

            if (flag.QuotePrefix)
            {
                target.QuotePrefix = style.QuotePrefix;
            }

            SetStyle(target);
        }

        private void AssignValue(object value)
        {
            if (value == null)
            {
                ClearValue();
                return;
            }

            if (value is string)
            {
                PutValue((string)value);
                return;
            }
            if (value is bool)
            {
                PutValue((bool)value);
                return;
            }
            if (value is DateTime)
            {
                PutValue((DateTime)value);
                return;
            }
            if (value is byte)
            {
                SetScalar((byte)value, CellValueKind.Number);
                return;
            }
            if (value is short)
            {
                SetScalar((short)value, CellValueKind.Number);
                return;
            }
            if (value is int)
            {
                PutValue((int)value);
                return;
            }
            if (value is long)
            {
                SetScalar((long)value, CellValueKind.Number);
                return;
            }
            if (value is float)
            {
                SetScalar((float)value, CellValueKind.Number);
                return;
            }
            if (value is double)
            {
                PutValue((double)value);
                return;
            }
            if (value is decimal)
            {
                SetScalar((decimal)value, CellValueKind.Number);
                return;
            }
            if (value is char)
            {
                PutValue(((char)value).ToString());
                return;
            }
            if (value is IFormattable)
            {
                SetScalar(value, CellValueKind.Number);
                return;
            }

            PutValue(value.ToString() ?? string.Empty);
        }

        private void ClearValue()
        {
            var record = GetOrCreateRecord();
            record.Value = null;
            record.Formula = null;
            record.Kind = CellValueKind.Blank;
        }

        private CellRecord TryGetRecord()
        {
            CellRecord record;
            _worksheet.Model.Cells.TryGetValue(_address, out record);
            return record;
        }

        private CellRecord GetOrCreateRecord()
        {
            CellRecord existing;
            if (_worksheet.Model.Cells.TryGetValue(_address, out existing))
            {
                return existing;
            }

            // Allocate a record even for currently blank cells so style-only and
            // formula-ready cells can participate in later serialization.
            var record = new CellRecord
            {
                Style = _worksheet.Workbook.Model.DefaultStyle.Clone(),
            };
            _worksheet.Model.Cells[_address] = record;
            return record;
        }

        private void SetScalar(object value, CellValueKind kind)
        {
            var record = GetOrCreateRecord();
            record.Value = value;
            record.Kind = kind;
            record.Formula = null;
        }

        private static string NormalizeFormula(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value[0] == '=' ? value.Substring(1) : value;
        }
    }
}
