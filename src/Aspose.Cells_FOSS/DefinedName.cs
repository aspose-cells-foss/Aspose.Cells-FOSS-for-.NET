using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents defined name.
    /// </summary>
    public sealed class DefinedName
    {
        private readonly Workbook _workbook;
        private readonly DefinedNameModel _model;

        internal DefinedName(Workbook workbook, DefinedNameModel model)
        {
            _workbook = workbook;
            _model = model;
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the formula.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the local sheet index.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether hidden.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the comment.
        /// </summary>
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
}
