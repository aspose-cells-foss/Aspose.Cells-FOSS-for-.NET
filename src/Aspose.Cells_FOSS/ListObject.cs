using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents an Excel table (structured reference / ListObject).
    /// </summary>
    public sealed class ListObject
    {
        private readonly ListObjectModel _model;
        private readonly WorksheetModel _worksheetModel;
        private readonly IListObjectOwner _owner;
        private readonly ListColumnCollection _listColumns;

        internal ListObject(ListObjectModel model, WorksheetModel worksheetModel, IListObjectOwner owner)
        {
            _model = model;
            _worksheetModel = worksheetModel;
            _owner = owner;
            _listColumns = new ListColumnCollection(model);
        }

        /// <summary>
        /// Gets or sets the user-visible table name. Must be non-empty and contain no spaces.
        /// </summary>
        public string DisplayName
        {
            get
            {
                return _model.DisplayName;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new CellsException("Table DisplayName must be non-empty.");
                }

                ListObjectSupport.ValidateDisplayName(value);
                _owner.ValidateUniqueDisplayName(value, _model);
                _model.DisplayName = value;
                _model.Name = value;
            }
        }

        /// <summary>
        /// Gets or sets an optional comment for the table.
        /// </summary>
        public string Comment
        {
            get
            {
                return _model.Comment;
            }
            set
            {
                _model.Comment = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets the zero-based row index of the first row (header or first data row).
        /// </summary>
        public int StartRow
        {
            get
            {
                return _model.StartRow;
            }
        }

        /// <summary>
        /// Gets the zero-based column index of the first column.
        /// </summary>
        public int StartColumn
        {
            get
            {
                return _model.StartColumn;
            }
        }

        /// <summary>
        /// Gets the zero-based row index of the last row (last data or totals row).
        /// </summary>
        public int EndRow
        {
            get
            {
                return _model.EndRow;
            }
        }

        /// <summary>
        /// Gets the zero-based column index of the last column.
        /// </summary>
        public int EndColumn
        {
            get
            {
                return _model.EndColumn;
            }
        }

        /// <summary>
        /// Gets or sets whether the first row of the table range is a header row.
        /// </summary>
        public bool ShowHeaderRow
        {
            get
            {
                return _model.ShowHeaderRow;
            }
            set
            {
                _model.ShowHeaderRow = value;
                _model.HasAutoFilter = value;
            }
        }

        /// <summary>
        /// Gets or sets whether the last row of the table range is a totals row.
        /// </summary>
        public bool ShowTotals
        {
            get
            {
                return _model.ShowTotals;
            }
            set
            {
                if (value == _model.ShowTotals)
                {
                    return;
                }

                if (value)
                {
                    _model.EndRow = _model.EndRow + 1;
                }
                else
                {
                    if (_model.EndRow > _model.StartRow)
                    {
                        _model.EndRow = _model.EndRow - 1;
                    }
                }

                _model.ShowTotals = value;
            }
        }

        /// <summary>
        /// Gets or sets the built-in table style type.
        /// Setting Custom preserves the current TableStyleName.
        /// </summary>
        public TableStyleType TableStyleType
        {
            get
            {
                return ListObjectSupport.TableStyleTypeFromName(_model.TableStyleName);
            }
            set
            {
                if (value != TableStyleType.Custom)
                {
                    _model.TableStyleName = ListObjectSupport.TableStyleTypeToName(value);
                }
            }
        }

        /// <summary>
        /// Gets or sets the raw table style name used in the SpreadsheetML tableStyleInfo element.
        /// Setting this to a built-in name also updates TableStyleType.
        /// </summary>
        public string TableStyleName
        {
            get
            {
                return _model.TableStyleName;
            }
            set
            {
                _model.TableStyleName = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets whether the first column receives banding or highlight formatting.
        /// </summary>
        public bool ShowTableStyleFirstColumn
        {
            get
            {
                return _model.ShowFirstColumn;
            }
            set
            {
                _model.ShowFirstColumn = value;
            }
        }

        /// <summary>
        /// Gets or sets whether the last column receives banding or highlight formatting.
        /// </summary>
        public bool ShowTableStyleLastColumn
        {
            get
            {
                return _model.ShowLastColumn;
            }
            set
            {
                _model.ShowLastColumn = value;
            }
        }

        /// <summary>
        /// Gets or sets whether alternating row stripes are shown.
        /// </summary>
        public bool ShowTableStyleRowStripes
        {
            get
            {
                return _model.ShowRowStripes;
            }
            set
            {
                _model.ShowRowStripes = value;
            }
        }

        /// <summary>
        /// Gets or sets whether alternating column stripes are shown.
        /// </summary>
        public bool ShowTableStyleColumnStripes
        {
            get
            {
                return _model.ShowColumnStripes;
            }
            set
            {
                _model.ShowColumnStripes = value;
            }
        }

        /// <summary>
        /// Gets the collection of columns in this table.
        /// </summary>
        public ListColumnCollection ListColumns
        {
            get
            {
                return _listColumns;
            }
        }

        /// <summary>
        /// Resizes the table to the specified range, rebuilding columns from header cells.
        /// </summary>
        public void Resize(int startRow, int startColumn, int endRow, int endColumn, bool hasHeaders)
        {
            ListObjectSupport.ValidateRange(startRow, startColumn, endRow, endColumn);
            _owner.ValidateNoOverlap(startRow, startColumn, endRow, endColumn, _model);
            _model.StartRow = startRow;
            _model.StartColumn = startColumn;
            _model.EndRow = endRow;
            _model.EndColumn = endColumn;
            _model.ShowHeaderRow = hasHeaders;
            _model.HasAutoFilter = hasHeaders;
            ListObjectSupport.RebuildColumns(_model, _worksheetModel);
        }

        /// <summary>
        /// Shows the autoFilter drop-down buttons on the table header row.
        /// </summary>
        public void ShowAutoFilter()
        {
            _model.HasAutoFilter = true;
        }

        /// <summary>
        /// Hides the autoFilter drop-down buttons from the table header row.
        /// </summary>
        public void RemoveAutoFilter()
        {
            _model.HasAutoFilter = false;
        }

        /// <summary>
        /// Removes the table structure, leaving the cell data in place.
        /// </summary>
        public void ConvertToRange()
        {
            _owner.RemoveTable(_model);
        }
    }
}
