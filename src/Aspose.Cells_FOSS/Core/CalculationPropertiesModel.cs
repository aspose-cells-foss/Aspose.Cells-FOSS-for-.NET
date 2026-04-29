using System.IO;
using System.Collections.Generic;
using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents calculation properties model.
    /// </summary>
    public sealed class CalculationPropertiesModel
    {
        /// <summary>
        /// Gets or sets the calculation id.
        /// </summary>
        public int? CalculationId { get; set; }
        /// <summary>
        /// Gets or sets the calculation mode.
        /// </summary>
        public string CalculationMode { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether full calculation on load.
        /// </summary>
        public bool FullCalculationOnLoad { get; set; }
        /// <summary>
        /// Gets or sets the reference mode.
        /// </summary>
        public string ReferenceMode { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether iterate.
        /// </summary>
        public bool Iterate { get; set; }
        /// <summary>
        /// Gets or sets the iterate count.
        /// </summary>
        public int? IterateCount { get; set; }
        /// <summary>
        /// Gets or sets the iterate delta.
        /// </summary>
        public double? IterateDelta { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether full precision.
        /// </summary>
        public bool? FullPrecision { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether calculation completed.
        /// </summary>
        public bool? CalculationCompleted { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether calculation on save.
        /// </summary>
        public bool? CalculationOnSave { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether concurrent calculation.
        /// </summary>
        public bool? ConcurrentCalculation { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether force full calculation.
        /// </summary>
        public bool ForceFullCalculation { get; set; }

        /// <summary>
        /// Copies values from the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        public void CopyFrom(CalculationPropertiesModel source)
        {
            CalculationId = source.CalculationId;
            CalculationMode = source.CalculationMode;
            FullCalculationOnLoad = source.FullCalculationOnLoad;
            ReferenceMode = source.ReferenceMode;
            Iterate = source.Iterate;
            IterateCount = source.IterateCount;
            IterateDelta = source.IterateDelta;
            FullPrecision = source.FullPrecision;
            CalculationCompleted = source.CalculationCompleted;
            CalculationOnSave = source.CalculationOnSave;
            ConcurrentCalculation = source.ConcurrentCalculation;
            ForceFullCalculation = source.ForceFullCalculation;
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState()
        {
            return CalculationId.HasValue
                || !string.IsNullOrEmpty(CalculationMode)
                || FullCalculationOnLoad
                || !string.IsNullOrEmpty(ReferenceMode)
                || Iterate
                || IterateCount.HasValue
                || IterateDelta.HasValue
                || FullPrecision.HasValue
                || CalculationCompleted.HasValue
                || CalculationOnSave.HasValue
                || ConcurrentCalculation.HasValue
                || ForceFullCalculation;
        }
    }
}
