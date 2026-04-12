using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

/// <summary>
/// Represents calculation properties.
/// </summary>
public sealed class CalculationProperties
{
    private readonly CalculationPropertiesModel _model;

    internal CalculationProperties(CalculationPropertiesModel model)
    {
        _model = model;
    }

    /// <summary>
    /// Gets or sets the calculation id.
    /// </summary>
    public int? CalculationId
    {
        get
        {
            return _model.CalculationId;
        }
        set
        {
            if (value.HasValue && value.Value < 0)
            {
                throw new CellsException("CalculationId must be non-negative.");
            }

            _model.CalculationId = value;
        }
    }

    /// <summary>
    /// Gets or sets the calculation mode.
    /// </summary>
    public string CalculationMode
    {
        get
        {
            return string.IsNullOrEmpty(_model.CalculationMode) ? "auto" : _model.CalculationMode;
        }
        set
        {
            _model.CalculationMode = WorkbookPropertySupport.NormalizeCalculationMode(value);
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether full calculation on load.
    /// </summary>
    public bool FullCalculationOnLoad
    {
        get
        {
            return _model.FullCalculationOnLoad;
        }
        set
        {
            _model.FullCalculationOnLoad = value;
        }
    }

    /// <summary>
    /// Gets or sets the reference mode.
    /// </summary>
    public string ReferenceMode
    {
        get
        {
            return string.IsNullOrEmpty(_model.ReferenceMode) ? "A1" : _model.ReferenceMode;
        }
        set
        {
            _model.ReferenceMode = WorkbookPropertySupport.NormalizeReferenceMode(value);
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether iterate.
    /// </summary>
    public bool Iterate
    {
        get
        {
            return _model.Iterate;
        }
        set
        {
            _model.Iterate = value;
        }
    }

    /// <summary>
    /// Gets or sets the iterate count.
    /// </summary>
    public int IterateCount
    {
        get
        {
            return _model.IterateCount ?? 100;
        }
        set
        {
            if (value < 0)
            {
                throw new CellsException("IterateCount must be non-negative.");
            }

            _model.IterateCount = value;
        }
    }

    /// <summary>
    /// Gets or sets the iterate delta.
    /// </summary>
    public double IterateDelta
    {
        get
        {
            return _model.IterateDelta ?? 0.001d;
        }
        set
        {
            if (value < 0d)
            {
                throw new CellsException("IterateDelta must be non-negative.");
            }

            _model.IterateDelta = value;
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether full precision.
    /// </summary>
    public bool FullPrecision
    {
        get
        {
            return !_model.FullPrecision.HasValue || _model.FullPrecision.Value;
        }
        set
        {
            _model.FullPrecision = value;
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether calculation completed.
    /// </summary>
    public bool CalculationCompleted
    {
        get
        {
            return !_model.CalculationCompleted.HasValue || _model.CalculationCompleted.Value;
        }
        set
        {
            _model.CalculationCompleted = value;
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether calculation on save.
    /// </summary>
    public bool CalculationOnSave
    {
        get
        {
            return !_model.CalculationOnSave.HasValue || _model.CalculationOnSave.Value;
        }
        set
        {
            _model.CalculationOnSave = value;
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether concurrent calculation.
    /// </summary>
    public bool ConcurrentCalculation
    {
        get
        {
            return !_model.ConcurrentCalculation.HasValue || _model.ConcurrentCalculation.Value;
        }
        set
        {
            _model.ConcurrentCalculation = value;
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether force full calculation.
    /// </summary>
    public bool ForceFullCalculation
    {
        get
        {
            return _model.ForceFullCalculation;
        }
        set
        {
            _model.ForceFullCalculation = value;
        }
    }
}
