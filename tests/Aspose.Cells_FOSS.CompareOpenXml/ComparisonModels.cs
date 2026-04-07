using Aspose.Cells_FOSS.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Aspose.Cells_FOSS.CompareOpenXml;

internal sealed record OpenXmlReadContext(WorkbookPart WorkbookPart, IReadOnlyList<string> SharedStrings, DateSystem DateSystem, OpenXmlStylesContext Styles);

internal sealed record OpenXmlCellDescriptor(string SheetName, string CellReference, DocumentFormat.OpenXml.Spreadsheet.Cell Cell);

internal sealed record OpenXmlStylesContext(IReadOnlyList<StyleSnapshot> CellFormats, ISet<int> DateStyleIndexes)
{
    public static OpenXmlStylesContext Default { get; } = new(new[] { StyleSnapshot.Default }, new HashSet<int>());
}

internal sealed record OpenXmlFontDescriptor(string Name, string Size, bool Bold, bool Italic, bool Underline, bool StrikeThrough, string Color)
{
    public static OpenXmlFontDescriptor Default { get; } = new("Calibri", ComparisonValueHelpers.NormalizeDouble(11d), false, false, false, false, ComparisonValueHelpers.DefaultColorHex);
}

internal sealed record OpenXmlFillDescriptor(string Pattern, string ForegroundColor, string BackgroundColor)
{
    public static OpenXmlFillDescriptor Default { get; } = new("None", ComparisonValueHelpers.DefaultColorHex, ComparisonValueHelpers.DefaultColorHex);
}

internal sealed record OpenXmlBorderSideDescriptor(string Style, string Color)
{
    public static OpenXmlBorderSideDescriptor Default { get; } = new("None", ComparisonValueHelpers.DefaultColorHex);
}

internal sealed record OpenXmlBorderDescriptor(OpenXmlBorderSideDescriptor Left, OpenXmlBorderSideDescriptor Right, OpenXmlBorderSideDescriptor Top, OpenXmlBorderSideDescriptor Bottom, OpenXmlBorderSideDescriptor Diagonal, bool DiagonalUp, bool DiagonalDown)
{
    public static OpenXmlBorderDescriptor Default { get; } = new(OpenXmlBorderSideDescriptor.Default, OpenXmlBorderSideDescriptor.Default, OpenXmlBorderSideDescriptor.Default, OpenXmlBorderSideDescriptor.Default, OpenXmlBorderSideDescriptor.Default, false, false);
}

internal sealed record CellSnapshot(string SheetName, string CellReference, string CellType, string Value, string Formula)
{
    public override string ToString()
    {
        return $"Sheet={SheetName}; Cell={CellReference}; Type={CellType}; Formula={Formula}; Value={Value}";
    }
}

internal sealed record StyleSnapshot(
    string SheetName,
    string CellReference,
    string FontName,
    string FontSize,
    bool FontBold,
    bool FontItalic,
    bool FontUnderline,
    bool FontStrikeThrough,
    string FontColor,
    string FillPattern,
    string FillForegroundColor,
    string FillBackgroundColor,
    string LeftBorderStyle,
    string LeftBorderColor,
    string RightBorderStyle,
    string RightBorderColor,
    string TopBorderStyle,
    string TopBorderColor,
    string BottomBorderStyle,
    string BottomBorderColor,
    string DiagonalBorderStyle,
    string DiagonalBorderColor,
    bool DiagonalUp,
    bool DiagonalDown,
    string HorizontalAlignment,
    string VerticalAlignment,
    bool WrapText,
    int IndentLevel,
    int TextRotation,
    bool ShrinkToFit,
    int ReadingOrder,
    int RelativeIndent,
    int NumberFormatId,
    string NumberFormatCode,
    bool IsLocked,
    bool IsHidden)
{
    public static StyleSnapshot Default { get; } = new(
        string.Empty,
        string.Empty,
        "Calibri",
        ComparisonValueHelpers.NormalizeDouble(11d),
        false,
        false,
        false,
        false,
        ComparisonValueHelpers.DefaultColorHex,
        "None",
        ComparisonValueHelpers.DefaultColorHex,
        ComparisonValueHelpers.DefaultColorHex,
        "None",
        ComparisonValueHelpers.DefaultColorHex,
        "None",
        ComparisonValueHelpers.DefaultColorHex,
        "None",
        ComparisonValueHelpers.DefaultColorHex,
        "None",
        ComparisonValueHelpers.DefaultColorHex,
        "None",
        ComparisonValueHelpers.DefaultColorHex,
        false,
        false,
        "General",
        "Bottom",
        false,
        0,
        0,
        false,
        0,
        0,
        0,
        string.Empty,
        true,
        false);

    public override string ToString()
    {
        return $"Sheet={SheetName}; Cell={CellReference}; Font={FontName}/{FontSize}/B={FontBold}/I={FontItalic}/U={FontUnderline}/S={FontStrikeThrough}/Color={FontColor}; Borders=L({LeftBorderStyle},{LeftBorderColor})/R({RightBorderStyle},{RightBorderColor})/T({TopBorderStyle},{TopBorderColor})/B({BottomBorderStyle},{BottomBorderColor})/D({DiagonalBorderStyle},{DiagonalBorderColor},Up={DiagonalUp},Down={DiagonalDown}); Align=H:{HorizontalAlignment}/V:{VerticalAlignment}/Wrap={WrapText}/Indent={IndentLevel}/Rotation={TextRotation}/Shrink={ShrinkToFit}/Reading={ReadingOrder}/RelativeIndent={RelativeIndent}; Fill={FillPattern}/{FillForegroundColor}/{FillBackgroundColor}; Number={NumberFormatId}/{NumberFormatCode}; Protection=Locked:{IsLocked}/Hidden:{IsHidden}";
    }
}

internal sealed record SnapshotMismatch(string CellKey, string LibrarySnapshot, string OpenXmlSnapshot);
