using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents collection of charts on a worksheet.
    /// </summary>
    /// <remarks>
    /// Charts provide visual data representation and support various chart types including bar, line, pie, etc.
    /// Charts are positioned by specifying anchor coordinates and reference worksheet data using formulas.
    /// Each chart can be independently configured with different types and styles.
    /// </remarks>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// var worksheet = workbook.Worksheets[0];
    ///
    /// // Prepare data
    /// worksheet.Cells["A1"].PutValue("Month");
    /// worksheet.Cells["B1"].PutValue("Sales");
    /// // ... add more data ...
    ///
    /// // Create a bar chart
    /// int chartIndex = worksheet.Charts.Add(
    ///     ChartType.Column,
    ///     "Sheet1!$B$1:$B$12",  // data range formula
    ///     5, 1, 20, 5);     // anchor coordinates
    ///
    /// Console.WriteLine($"Chart {chartIndex} added: {worksheet.Charts[chartIndex].Name}");
    /// </code>
    /// </example>
    public sealed class ChartCollection
    {
        private readonly WorksheetModel _model;

        internal ChartCollection(WorksheetModel model)
        {
            _model = model;
        }

        /// <summary>Gets the number of charts on the worksheet.</summary>
        public int Count => _model.Charts.Count;

        /// <summary>Gets the chart at the specified zero-based index.</summary>
        public Chart this[int index]
        {
            get
            {
                if (index < 0 || index >= _model.Charts.Count)
                {
                    throw new CellsException("Chart index " + index + " is out of range. The worksheet contains " + _model.Charts.Count + " chart(s).");
                }

                return new Chart(_model.Charts[index]);
            }
        }

        /// <summary>
        /// Adds a new chart of the specified type to the worksheet.
        /// </summary>
        /// <param name="type">The chart type. Modern chartex types (Waterfall, Treemap, etc.) are not supported for programmatic creation.</param>
        /// <param name="dataRange">The cell range formula for the chart series data (e.g. "Sheet1!$B$1:$B$5").</param>
        /// <param name="upperLeftRow">Zero-based row index of the chart's upper-left anchor.</param>
        /// <param name="upperLeftColumn">Zero-based column index of the chart's upper-left anchor.</param>
        /// <param name="lowerRightRow">Zero-based row index of the chart's lower-right anchor.</param>
        /// <param name="lowerRightColumn">Zero-based column index of the chart's lower-right anchor.</param>
        /// <returns>The zero-based index of the added chart.</returns>
        public int Add(ChartType type, string dataRange,
                       int upperLeftRow, int upperLeftColumn,
                       int lowerRightRow, int lowerRightColumn)
        {
            var chartId = _model.Charts.Count + 1;
            var rawXml = ChartXmlTemplates.Build(type, dataRange, "Chart " + chartId);
            var model = new ChartModel
            {
                Name = "Chart " + chartId,
                ChartType = type,
                UpperLeftRow = upperLeftRow,
                UpperLeftColumn = upperLeftColumn,
                LowerRightRow = lowerRightRow,
                LowerRightColumn = lowerRightColumn,
                RawChartXml = rawXml,
            };
            _model.Charts.Add(model);
            return _model.Charts.Count - 1;
        }
    }
}
