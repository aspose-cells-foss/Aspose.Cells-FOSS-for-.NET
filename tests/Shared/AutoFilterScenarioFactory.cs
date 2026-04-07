using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Testing;

public static class AutoFilterScenarioFactory
{
    public static Workbook CreateAutoFilterWorkbook()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "Filtered";

        sheet.Cells[0, 0].PutValue("Status");
        sheet.Cells[0, 1].PutValue("Amount");
        sheet.Cells[0, 2].PutValue("Color");
        sheet.Cells[0, 3].PutValue("Date");
        sheet.Cells[0, 4].PutValue("Score");
        sheet.Cells[1, 0].PutValue("Open");
        sheet.Cells[1, 1].PutValue(10);
        sheet.Cells[1, 2].PutValue("Red");
        sheet.Cells[1, 3].PutValue(new DateTime(2024, 5, 1));
        sheet.Cells[1, 4].PutValue(70);
        sheet.Cells[2, 0].PutValue("Closed");
        sheet.Cells[2, 1].PutValue(20);
        sheet.Cells[2, 2].PutValue("Blue");
        sheet.Cells[2, 3].PutValue(new DateTime(2024, 5, 2));
        sheet.Cells[2, 4].PutValue(80);
        sheet.Cells[3, 0].PutValue("Open");
        sheet.Cells[3, 1].PutValue(30);
        sheet.Cells[3, 2].PutValue("Green");
        sheet.Cells[3, 3].PutValue(new DateTime(2024, 5, 3));
        sheet.Cells[3, 4].PutValue(90);
        sheet.Cells[4, 0].PutValue("Pending");
        sheet.Cells[4, 1].PutValue(40);
        sheet.Cells[4, 2].PutValue("Yellow");
        sheet.Cells[4, 3].PutValue(new DateTime(2024, 5, 4));
        sheet.Cells[4, 4].PutValue(60);
        sheet.Cells[5, 0].PutValue("Closed");
        sheet.Cells[5, 1].PutValue(50);
        sheet.Cells[5, 2].PutValue("Black");
        sheet.Cells[5, 3].PutValue(new DateTime(2024, 5, 5));
        sheet.Cells[5, 4].PutValue(50);
        AddDifferentialStyles(sheet);

        sheet.AutoFilter.Range = "A1:E6";

        var statusColumn = sheet.AutoFilter.FilterColumns[sheet.AutoFilter.FilterColumns.Add(0)];
        statusColumn.HiddenButton = true;
        statusColumn.Filters.Add("Open");
        statusColumn.Filters.Add("Closed");

        var amountColumn = sheet.AutoFilter.FilterColumns[sheet.AutoFilter.FilterColumns.Add(1)];
        amountColumn.CustomFilters.MatchAll = true;
        amountColumn.CustomFilters.Add(FilterOperatorType.GreaterOrEqual, "10");
        amountColumn.CustomFilters.Add(FilterOperatorType.LessOrEqual, "50");

        var colorColumn = sheet.AutoFilter.FilterColumns[sheet.AutoFilter.FilterColumns.Add(2)];
        colorColumn.ColorFilter.Enabled = true;
        colorColumn.ColorFilter.DifferentialStyleId = 3;
        colorColumn.ColorFilter.CellColor = true;

        var dateColumn = sheet.AutoFilter.FilterColumns[sheet.AutoFilter.FilterColumns.Add(3)];
        dateColumn.DynamicFilter.Enabled = true;
        dateColumn.DynamicFilter.Type = "thisMonth";
        dateColumn.DynamicFilter.Value = 1d;
        dateColumn.DynamicFilter.MaxValue = 31d;

        var scoreColumn = sheet.AutoFilter.FilterColumns[sheet.AutoFilter.FilterColumns.Add(4)];
        scoreColumn.Top10.Enabled = true;
        scoreColumn.Top10.Top = false;
        scoreColumn.Top10.Percent = true;
        scoreColumn.Top10.Value = 10d;
        scoreColumn.Top10.FilterValue = 2.5d;

        sheet.AutoFilter.SortState.Ref = "A2:E6";
        sheet.AutoFilter.SortState.CaseSensitive = true;
        sheet.AutoFilter.SortState.SortMethod = "pinYin";

        var valueSort = sheet.AutoFilter.SortState.SortConditions[sheet.AutoFilter.SortState.SortConditions.Add("B2:B6")];
        valueSort.Descending = true;
        valueSort.SortBy = "value";
        valueSort.CustomList = "High,Medium,Low";

        var colorSort = sheet.AutoFilter.SortState.SortConditions[sheet.AutoFilter.SortState.SortConditions.Add("C2:C6")];
        colorSort.SortBy = "cellColor";
        colorSort.DifferentialStyleId = 4;

        var iconSort = sheet.AutoFilter.SortState.SortConditions[sheet.AutoFilter.SortState.SortConditions.Add("E2:E6")];
        iconSort.SortBy = "icon";
        iconSort.IconSet = "3TrafficLights1";
        iconSort.IconId = 2;

        return workbook;
    }

    public static void AssertAutoFilter(Workbook workbook)
    {
        var sheet = workbook.Worksheets[0];
        AssertEx.Equal("A1:E6", sheet.AutoFilter.Range);
        AssertEx.Equal(5, sheet.AutoFilter.FilterColumns.Count);

        var statusColumn = sheet.AutoFilter.FilterColumns[0];
        AssertEx.Equal(0, statusColumn.ColumnIndex);
        AssertEx.True(statusColumn.HiddenButton);
        AssertEx.Equal(2, statusColumn.Filters.Count);
        AssertEx.Equal("Open", statusColumn.Filters[0]);
        AssertEx.Equal("Closed", statusColumn.Filters[1]);

        var amountColumn = sheet.AutoFilter.FilterColumns[1];
        AssertEx.Equal(1, amountColumn.ColumnIndex);
        AssertEx.True(amountColumn.CustomFilters.MatchAll);
        AssertEx.Equal(2, amountColumn.CustomFilters.Count);
        AssertEx.Equal(FilterOperatorType.GreaterOrEqual, amountColumn.CustomFilters[0].Operator);
        AssertEx.Equal("10", amountColumn.CustomFilters[0].Value);
        AssertEx.Equal(FilterOperatorType.LessOrEqual, amountColumn.CustomFilters[1].Operator);
        AssertEx.Equal("50", amountColumn.CustomFilters[1].Value);

        var colorColumn = sheet.AutoFilter.FilterColumns[2];
        AssertEx.Equal(2, colorColumn.ColumnIndex);
        AssertEx.True(colorColumn.ColorFilter.Enabled);
        AssertEx.Equal(3, colorColumn.ColorFilter.DifferentialStyleId ?? -1);
        AssertEx.True(colorColumn.ColorFilter.CellColor);

        var dateColumn = sheet.AutoFilter.FilterColumns[3];
        AssertEx.Equal(3, dateColumn.ColumnIndex);
        AssertEx.True(dateColumn.DynamicFilter.Enabled);
        AssertEx.Equal("thisMonth", dateColumn.DynamicFilter.Type);
        AssertEx.Equal(1d, dateColumn.DynamicFilter.Value ?? 0d);
        AssertEx.Equal(31d, dateColumn.DynamicFilter.MaxValue ?? 0d);

        var scoreColumn = sheet.AutoFilter.FilterColumns[4];
        AssertEx.Equal(4, scoreColumn.ColumnIndex);
        AssertEx.True(scoreColumn.Top10.Enabled);
        AssertEx.False(scoreColumn.Top10.Top);
        AssertEx.True(scoreColumn.Top10.Percent);
        AssertEx.Equal(10d, scoreColumn.Top10.Value ?? 0d);
        AssertEx.Equal(2.5d, scoreColumn.Top10.FilterValue ?? 0d);

        AssertEx.Equal("A2:E6", sheet.AutoFilter.SortState.Ref);
        AssertEx.True(sheet.AutoFilter.SortState.CaseSensitive);
        AssertEx.Equal("pinYin", sheet.AutoFilter.SortState.SortMethod);
        AssertEx.Equal(3, sheet.AutoFilter.SortState.SortConditions.Count);

        var valueSort = sheet.AutoFilter.SortState.SortConditions[0];
        AssertEx.Equal("B2:B6", valueSort.Ref);
        AssertEx.True(valueSort.Descending);
        AssertEx.Equal("value", valueSort.SortBy);
        AssertEx.Equal("High,Medium,Low", valueSort.CustomList);
        AssertEx.Null(valueSort.DifferentialStyleId);
        AssertEx.Equal(string.Empty, valueSort.IconSet);
        AssertEx.Null(valueSort.IconId);

        var colorSort = sheet.AutoFilter.SortState.SortConditions[1];
        AssertEx.Equal("C2:C6", colorSort.Ref);
        AssertEx.False(colorSort.Descending);
        AssertEx.Equal("cellColor", colorSort.SortBy);
        AssertEx.Equal(string.Empty, colorSort.CustomList);
        AssertEx.Equal(4, colorSort.DifferentialStyleId ?? -1);
        AssertEx.Equal(string.Empty, colorSort.IconSet);
        AssertEx.Null(colorSort.IconId);

        var iconSort = sheet.AutoFilter.SortState.SortConditions[2];
        AssertEx.Equal("E2:E6", iconSort.Ref);
        AssertEx.False(iconSort.Descending);
        AssertEx.Equal("icon", iconSort.SortBy);
        AssertEx.Equal(string.Empty, iconSort.CustomList);
        AssertEx.Null(iconSort.DifferentialStyleId);
        AssertEx.Equal("3TrafficLights1", iconSort.IconSet);
        AssertEx.Equal(2, iconSort.IconId ?? -1);
    }

    private static void AddDifferentialStyles(Worksheet sheet)
    {
        var colors = new[]
        {
            Color.FromArgb(255, 255, 235, 156),
            Color.FromArgb(255, 198, 239, 206),
            Color.FromArgb(255, 221, 235, 247),
            Color.FromArgb(255, 255, 199, 206),
            Color.FromArgb(255, 189, 215, 238),
        };

        for (var index = 0; index < colors.Length; index++)
        {
            var collection = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
            collection.AddArea(CellArea.CreateCellArea(index + 1, 0, index + 1, 0));
            var condition = collection[collection.AddCondition(FormatConditionType.Expression, OperatorType.None, "TRUE", string.Empty)];
            var style = condition.Style;
            style.Pattern = FillPattern.Solid;
            style.ForegroundColor = colors[index];
            condition.Style = style;
        }
    }
}