using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a chart embedded in a worksheet.
    /// </summary>
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
