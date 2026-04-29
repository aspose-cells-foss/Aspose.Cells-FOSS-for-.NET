using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents worksheet protection.
    /// </summary>
    public sealed class WorksheetProtection
    {
        private readonly WorksheetProtectionModel _model;

        internal WorksheetProtection(WorksheetProtectionModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets a value indicating whether protected.
        /// </summary>
        public bool IsProtected
        {
            get
            {
                return _model.IsProtected;
            }
            set
            {
                _model.IsProtected = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether objects.
        /// </summary>
        public bool Objects
        {
            get
            {
                return _model.Objects;
            }
            set
            {
                _model.Objects = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether scenarios.
        /// </summary>
        public bool Scenarios
        {
            get
            {
                return _model.Scenarios;
            }
            set
            {
                _model.Scenarios = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether format cells.
        /// </summary>
        public bool FormatCells
        {
            get
            {
                return _model.FormatCells;
            }
            set
            {
                _model.FormatCells = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether format columns.
        /// </summary>
        public bool FormatColumns
        {
            get
            {
                return _model.FormatColumns;
            }
            set
            {
                _model.FormatColumns = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether format rows.
        /// </summary>
        public bool FormatRows
        {
            get
            {
                return _model.FormatRows;
            }
            set
            {
                _model.FormatRows = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether insert columns.
        /// </summary>
        public bool InsertColumns
        {
            get
            {
                return _model.InsertColumns;
            }
            set
            {
                _model.InsertColumns = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether insert rows.
        /// </summary>
        public bool InsertRows
        {
            get
            {
                return _model.InsertRows;
            }
            set
            {
                _model.InsertRows = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether insert hyperlinks.
        /// </summary>
        public bool InsertHyperlinks
        {
            get
            {
                return _model.InsertHyperlinks;
            }
            set
            {
                _model.InsertHyperlinks = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether delete columns.
        /// </summary>
        public bool DeleteColumns
        {
            get
            {
                return _model.DeleteColumns;
            }
            set
            {
                _model.DeleteColumns = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether delete rows.
        /// </summary>
        public bool DeleteRows
        {
            get
            {
                return _model.DeleteRows;
            }
            set
            {
                _model.DeleteRows = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether select locked cells.
        /// </summary>
        public bool SelectLockedCells
        {
            get
            {
                return _model.SelectLockedCells;
            }
            set
            {
                _model.SelectLockedCells = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether sort.
        /// </summary>
        public bool Sort
        {
            get
            {
                return _model.Sort;
            }
            set
            {
                _model.Sort = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether auto filter.
        /// </summary>
        public bool AutoFilter
        {
            get
            {
                return _model.AutoFilter;
            }
            set
            {
                _model.AutoFilter = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether pivot tables.
        /// </summary>
        public bool PivotTables
        {
            get
            {
                return _model.PivotTables;
            }
            set
            {
                _model.PivotTables = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether select unlocked cells.
        /// </summary>
        public bool SelectUnlockedCells
        {
            get
            {
                return _model.SelectUnlockedCells;
            }
            set
            {
                _model.SelectUnlockedCells = value;
                MarkProtectedWhenEnabled(value);
            }
        }

        internal void Reset()
        {
            _model.Clear();
        }

        private void MarkProtectedWhenEnabled(bool value)
        {
            if (value)
            {
                _model.IsProtected = true;
            }
        }
    }
}
