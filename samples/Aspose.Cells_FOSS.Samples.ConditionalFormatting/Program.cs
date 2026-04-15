using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.ConditionalFormatting
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "conditional-formatting-sample.xlsx");

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Conditional Formatting";

            for (var index = 0; index < 10; index++)
            {
                sheet.Cells[index, 0].PutValue(index + 1);
                sheet.Cells[index, 1].PutValue((index + 1) * 10);
                sheet.Cells[index, 2].PutValue((index + 1) * 10);
                sheet.Cells[index, 3].PutValue((index + 1) * 10);
                sheet.Cells[index, 4].PutValue((index + 1) * 10);
            }

            var betweenCollection = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
            betweenCollection.AddArea(CellArea.CreateCellArea("A1", "A10"));
            var betweenRule = betweenCollection[betweenCollection.AddCondition(FormatConditionType.CellValue, OperatorType.Between, "3", "7")];
            betweenRule.Priority = 1;
            betweenRule.StopIfTrue = true;
            var betweenStyle = betweenRule.Style;
            betweenStyle.Pattern = FillPattern.Solid;
            betweenStyle.ForegroundColor = Color.FromArgb(255, 255, 199, 206);
            betweenStyle.Font.Bold = true;
            betweenStyle.Font.Color = Color.FromArgb(255, 156, 0, 6);
            betweenRule.Style = betweenStyle;

            var expressionCollection = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
            expressionCollection.AddArea(CellArea.CreateCellArea("B1", "B10"));
            var expressionRule = expressionCollection[expressionCollection.AddCondition(FormatConditionType.Expression, OperatorType.None, "MOD(B1,20)=0", string.Empty)];
            expressionRule.Priority = 2;
            var expressionStyle = expressionRule.Style;
            expressionStyle.Font.Italic = true;
            expressionStyle.Font.Color = Color.FromArgb(255, 0, 0, 255);
            expressionRule.Style = expressionStyle;

            var colorScaleCollection = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
            colorScaleCollection.AddArea(CellArea.CreateCellArea("C1", "C10"));
            var colorScaleRule = colorScaleCollection[colorScaleCollection.AddCondition(FormatConditionType.ColorScale)];
            colorScaleRule.ColorScaleCount = 3;
            colorScaleRule.MinColor = Color.FromArgb(255, 248, 105, 107);
            colorScaleRule.MidColor = Color.FromArgb(255, 255, 235, 132);
            colorScaleRule.MaxColor = Color.FromArgb(255, 99, 190, 123);

            var dataBarCollection = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
            dataBarCollection.AddArea(CellArea.CreateCellArea("D1", "D10"));
            var dataBarRule = dataBarCollection[dataBarCollection.AddCondition(FormatConditionType.DataBar)];
            dataBarRule.BarColor = Color.FromArgb(255, 99, 142, 198);
            dataBarRule.NegativeBarColor = Color.FromArgb(255, 255, 0, 0);
            dataBarRule.ShowBorder = true;
            dataBarRule.Direction = "left-to-right";

            var iconSetCollection = sheet.ConditionalFormattings[sheet.ConditionalFormattings.Add()];
            iconSetCollection.AddArea(CellArea.CreateCellArea("E1", "E10"));
            var iconSetRule = iconSetCollection[iconSetCollection.AddCondition(FormatConditionType.IconSet)];
            iconSetRule.IconSetType = "4Arrows";
            iconSetRule.ReverseIcons = true;
            iconSetRule.ShowIconOnly = true;

            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedSheet = loaded.Worksheets["Conditional Formatting"];

            Console.WriteLine("Saved: " + outputPath);
            Console.WriteLine("Conditional formatting collections: " + loadedSheet.ConditionalFormattings.Count);
            Console.WriteLine("A-column rule type: " + loadedSheet.ConditionalFormattings[0][0].Type);
            Console.WriteLine("C-column rule type: " + loadedSheet.ConditionalFormattings[2][0].Type);
            Console.WriteLine("E-column icon set: " + loadedSheet.ConditionalFormattings[4][0].IconSetType);
        }
    }
}
