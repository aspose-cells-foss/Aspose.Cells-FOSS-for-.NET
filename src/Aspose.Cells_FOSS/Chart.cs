using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a chart embedded in a worksheet.
    /// </summary>
    /// <remarks>
    /// Charts provide visual representation of data and can be created programmatically or loaded from existing XLSX files.
    /// Note that modern chartex types (Waterfall, Treemap, Sunburst, etc.) must be loaded from existing files;
    /// they cannot be created using the <see cref="ChartCollection.Add"/> method.
    /// </remarks>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// var worksheet = workbook.Worksheets[0];
    ///
    /// // Add sample data
    /// worksheet.Cells["A1"].PutValue("Month");
    /// worksheet.Cells["B1"].PutValue("Sales");
    /// worksheet.Cells["A2"].PutValue("Jan");
    /// worksheet.Cells["B2"].PutValue(1000);
    /// worksheet.Cells["A3"].PutValue("Feb");
    /// worksheet.Cells["B3"].PutValue(1200);
    ///
    /// // Add a bar chart
    /// int chartIndex = worksheet.Charts.Add(
    ///     ChartType.Column,
    ///     "Sheet1!$B$1:$B$3",
    ///     5, 1, 10, 3);
    ///
    /// // Access and configure the chart
    /// var chart = worksheet.Charts[chartIndex];
    /// Console.WriteLine($"Chart: {chart.Name}");
    /// </code>
    /// </example>
    public sealed class Chart
    {
        private readonly ChartModel _model;

        internal Chart(ChartModel model)
        {
            _model = model;
        }

        /// <summary>Gets the chart display name.</summary>
        public string Name => _model.Name;

        /// <summary>Gets the chart type.</summary>
        public ChartType ChartType => _model.ChartType;

        /// <summary>Gets the zero-based upper-left row of the chart anchor.</summary>
        public int UpperLeftRow => _model.UpperLeftRow;

        /// <summary>Gets the zero-based upper-left column of the chart anchor.</summary>
        public int UpperLeftColumn => _model.UpperLeftColumn;

        /// <summary>Gets the zero-based lower-right row of the chart anchor.</summary>
        public int LowerRightRow => _model.LowerRightRow;

        /// <summary>Gets the zero-based lower-right column of the chart anchor.</summary>
        public int LowerRightColumn => _model.LowerRightColumn;

        /// <summary>Gets the width extent in EMU.</summary>
        public long ExtentCx => _model.ExtentCx;

        /// <summary>Gets the height extent in EMU.</summary>
        public long ExtentCy => _model.ExtentCy;
    }
}
