using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a single column in an Excel table.
    /// </summary>
    public sealed class ListColumn
    {
        private readonly ListColumnModel _model;

        internal ListColumn(ListColumnModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets the column header name displayed in the header row.
        /// </summary>
        public string Name
        {
            get
            {
                return _model.Name;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new CellsException("ListColumn name must be non-empty.");
                }

                _model.Name = value;
            }
        }

        /// <summary>
        /// Gets or sets the aggregation function shown in the totals row cell.
        /// </summary>
        public TotalsCalculation TotalsCalculation
        {
            get
            {
                return TotalsCalculationFromString(_model.TotalsRowFunction);
            }
            set
            {
                _model.TotalsRowFunction = TotalsCalculationToString(value);
            }
        }

        /// <summary>
        /// Gets or sets the static label shown in the totals row when TotalsCalculation is None.
        /// </summary>
        public string TotalsRowLabel
        {
            get
            {
                return _model.TotalsRowLabel;
            }
            set
            {
                _model.TotalsRowLabel = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the custom formula used when TotalsCalculation is Custom.
        /// </summary>
        public string TotalsRowFormula
        {
            get
            {
                return _model.TotalsRowFormula;
            }
            set
            {
                _model.TotalsRowFormula = value ?? string.Empty;
            }
        }

        internal static TotalsCalculation TotalsCalculationFromString(string value)
        {
            switch (value)
            {
                case "sum": return TotalsCalculation.Sum;
                case "count": return TotalsCalculation.Count;
                case "average": return TotalsCalculation.Average;
                case "max": return TotalsCalculation.Max;
                case "min": return TotalsCalculation.Min;
                case "stdDev": return TotalsCalculation.StdDev;
                case "var": return TotalsCalculation.Var;
                case "countNums": return TotalsCalculation.CountNums;
                case "custom": return TotalsCalculation.Custom;
                default: return TotalsCalculation.None;
            }
        }

        internal static string TotalsCalculationToString(TotalsCalculation value)
        {
            switch (value)
            {
                case TotalsCalculation.Sum: return "sum";
                case TotalsCalculation.Count: return "count";
                case TotalsCalculation.Average: return "average";
                case TotalsCalculation.Max: return "max";
                case TotalsCalculation.Min: return "min";
                case TotalsCalculation.StdDev: return "stdDev";
                case TotalsCalculation.Var: return "var";
                case TotalsCalculation.CountNums: return "countNums";
                case TotalsCalculation.Custom: return "custom";
                default: return "none";
            }
        }
    }
}
