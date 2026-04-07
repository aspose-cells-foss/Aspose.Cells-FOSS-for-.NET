using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public class Cell
{
    private readonly Worksheet _worksheet;
    private readonly CellAddress _address;

    internal Cell(Worksheet worksheet, CellAddress address)
    {
        _worksheet = worksheet;
        _address = address;
    }

    public object? Value
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

    public string StringValue
    {
        get
        {
            var record = TryGetRecord();
            return DisplayTextFormatter.FormatStringValue(record?.Value);
        }
    }

    public string DisplayStringValue
    {
        get
        {
            var record = TryGetRecord();
            var style = record is null ? _worksheet.Workbook.Model.DefaultStyle : record.Style;
            return DisplayTextFormatter.FormatDisplayValue(record?.Value, style, _worksheet.Workbook.Model.Settings.DisplayCulture);
        }
    }

    public string Formula
    {
        get
        {
            var record = TryGetRecord();
            if (record is null || string.IsNullOrEmpty(record.Formula))
            {
                return string.Empty;
            }

            return "=" + record.Formula;
        }
        set
        {
            var record = GetOrCreateRecord();
            record.Formula = NormalizeFormula(value);
            if (string.IsNullOrEmpty(record.Formula))
            {
                if (record.Value is null)
                {
                    record.Kind = CellValueKind.Blank;
                }

                return;
            }

            record.Kind = CellValueKind.Formula;
        }
    }

    public CellValueType Type
    {
        get
        {
            var record = TryGetRecord();
            if (record is null)
            {
                return CellValueType.Blank;
            }

            if (!string.IsNullOrEmpty(record.Formula))
            {
                return CellValueType.Formula;
            }

            switch (record.Kind)
            {
                case CellValueKind.String:
                    return CellValueType.String;
                case CellValueKind.Number:
                    return CellValueType.Number;
                case CellValueKind.Boolean:
                    return CellValueType.Boolean;
                case CellValueKind.DateTime:
                    return CellValueType.DateTime;
                case CellValueKind.Formula:
                    return CellValueType.Formula;
                default:
                    return CellValueType.Blank;
            }
        }
    }

    public void PutValue(string value)
    {
        if (value is null) throw new ArgumentNullException(nameof(value));
        SetScalar(value, CellValueKind.String);
    }

    public void PutValue(int value)
    {
        SetScalar(value, CellValueKind.Number);
    }

    public void PutValue(decimal value)
    {
        SetScalar(value, CellValueKind.Number);
    }

    public void PutValue(double value)
    {
        SetScalar(value, CellValueKind.Number);
    }

    public void PutValue(bool value)
    {
        SetScalar(value, CellValueKind.Boolean);
    }

    public void PutValue(DateTime value)
    {
        SetScalar(value, CellValueKind.DateTime);
    }

    public Style GetStyle()
    {
        var record = TryGetRecord();
        var style = record?.Style ?? _worksheet.Workbook.Model.DefaultStyle;
        return Style.FromCore(style.Clone());
    }

    public void SetStyle(Style style)
    {
        if (style is null) throw new ArgumentNullException(nameof(style));
        GetOrCreateRecord().Style = style.ToCore();
    }

    private void AssignValue(object? value)
    {
        if (value is null)
        {
            ClearValue();
            return;
        }

        switch (value)
        {
            case string text:
                PutValue(text);
                return;
            case bool booleanValue:
                PutValue(booleanValue);
                return;
            case DateTime dateTime:
                PutValue(dateTime);
                return;
            case byte byteValue:
                SetScalar(byteValue, CellValueKind.Number);
                return;
            case short shortValue:
                SetScalar(shortValue, CellValueKind.Number);
                return;
            case int intValue:
                PutValue(intValue);
                return;
            case long longValue:
                SetScalar(longValue, CellValueKind.Number);
                return;
            case float floatValue:
                SetScalar(floatValue, CellValueKind.Number);
                return;
            case double doubleValue:
                PutValue(doubleValue);
                return;
            case decimal decimalValue:
                PutValue(decimalValue);
                return;
            case char characterValue:
                PutValue(characterValue.ToString());
                return;
            default:
                if (value is IFormattable)
                {
                    SetScalar(value, CellValueKind.Number);
                    return;
                }

                PutValue(value.ToString() ?? string.Empty);
                return;
        }
    }

    private void ClearValue()
    {
        var record = GetOrCreateRecord();
        record.Value = null;
        record.Formula = null;
        record.Kind = CellValueKind.Blank;
    }

    private CellRecord? TryGetRecord()
    {
        _worksheet.Model.Cells.TryGetValue(_address, out var record);
        return record;
    }

    private CellRecord GetOrCreateRecord()
    {
        if (_worksheet.Model.Cells.TryGetValue(_address, out var existing))
        {
            return existing;
        }

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
