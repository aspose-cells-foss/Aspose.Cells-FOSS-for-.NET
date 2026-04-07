using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public sealed class DefinedName
{
    private readonly Workbook _workbook;
    private readonly DefinedNameModel _model;

    internal DefinedName(Workbook workbook, DefinedNameModel model)
    {
        _workbook = workbook;
        _model = model;
    }

    public string Name
    {
        get
        {
            return _model.Name;
        }
        set
        {
            var normalized = DefinedNameUtility.NormalizeName(value);
            _workbook.EnsureUniqueDefinedName(_model, normalized, _model.LocalSheetIndex);
            _model.Name = normalized;
        }
    }

    public string Formula
    {
        get
        {
            return _model.Formula;
        }
        set
        {
            _model.Formula = DefinedNameUtility.NormalizeFormula(value);
        }
    }

    public int? LocalSheetIndex
    {
        get
        {
            return _model.LocalSheetIndex;
        }
        set
        {
            _workbook.EnsureValidDefinedNameScope(value);
            _workbook.EnsureUniqueDefinedName(_model, _model.Name, value);
            _model.LocalSheetIndex = value;
        }
    }

    public bool Hidden
    {
        get
        {
            return _model.Hidden;
        }
        set
        {
            _model.Hidden = value;
        }
    }

    public string Comment
    {
        get
        {
            return _model.Comment;
        }
        set
        {
            _model.Comment = DefinedNameUtility.NormalizeComment(value);
        }
    }
}
