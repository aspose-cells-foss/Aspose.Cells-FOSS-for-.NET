using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents the collection of charts on a worksheet.
    /// </summary>
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
