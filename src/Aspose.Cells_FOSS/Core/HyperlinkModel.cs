namespace Aspose.Cells_FOSS.Core;

public sealed class HyperlinkModel
{
    public int FirstRow { get; set; }
    public int FirstColumn { get; set; }
    public int TotalRows { get; set; } = 1;
    public int TotalColumns { get; set; } = 1;
    public string? Address { get; set; }
    public string? SubAddress { get; set; }
    public string? ScreenTip { get; set; }
    public string? TextToDisplay { get; set; }
}