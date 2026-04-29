using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents row.
    /// </summary>
    public sealed class Row
    {
        private readonly Worksheet _worksheet;
        private readonly int _index;

        internal Row(Worksheet worksheet, int index)
        {
            _worksheet = worksheet;
            _index = index;
        }

        /// <summary>
        /// Gets or sets the height.
        /// </summary>
        public double? Height
        {
            get
            {
                var model = TryGetModel();
                return model?.Height;
            }
            set
            {
                if (value.HasValue && value.Value <= 0d)
                {
                    throw new CellsException("Row height must be positive.");
                }

                var model = GetOrCreateModel();
                model.Height = value;
                Normalize();
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether hidden.
        /// </summary>
        public bool IsHidden
        {
            get
            {
                var model = TryGetModel();
                return model?.Hidden ?? false;
            }
            set
            {
                var model = GetOrCreateModel();
                model.Hidden = value;
                Normalize();
            }
        }

        private RowModel TryGetModel()
        {
            RowModel model;
            _worksheet.Model.Rows.TryGetValue(_index, out model);
            return model;
        }

        private RowModel GetOrCreateModel()
        {
            RowModel existing;
            if (_worksheet.Model.Rows.TryGetValue(_index, out existing))
            {
                return existing;
            }

            var created = new RowModel();
            _worksheet.Model.Rows[_index] = created;
            return created;
        }

        private void Normalize()
        {
            RowModel model;
            if (_worksheet.Model.Rows.TryGetValue(_index, out model)
                && !model.Height.HasValue
                && !model.Hidden
                && !model.StyleIndex.HasValue)
            {
                _worksheet.Model.Rows.Remove(_index);
            }
        }
    }
}
