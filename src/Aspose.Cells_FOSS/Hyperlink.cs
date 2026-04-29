using System.IO;
using System;
using System.Collections.Generic;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents hyperlink.
    /// </summary>
    public sealed class Hyperlink
    {
        private readonly IList<HyperlinkModel> _owner;
        private readonly HyperlinkModel _model;

        internal Hyperlink(IList<HyperlinkModel> owner, HyperlinkModel model)
        {
            _owner = owner;
            _model = model;
        }

        /// <summary>
        /// Gets the area.
        /// </summary>
        public string Area
        {
            get
            {
                var first = new CellAddress(_model.FirstRow, _model.FirstColumn).ToString();
                if (_model.TotalRows == 1 && _model.TotalColumns == 1)
                {
                    return first;
                }

                var last = new CellAddress(_model.FirstRow + _model.TotalRows - 1, _model.FirstColumn + _model.TotalColumns - 1).ToString();
                return first + ":" + last;
            }
        }

        /// <summary>
        /// Gets or sets the address.
        /// </summary>
        public string Address
        {
            get
            {
                var address = _model.Address;
                if (!string.IsNullOrEmpty(address))
                {
                    return address;
                }

                return _model.SubAddress ?? string.Empty;
            }
            set
            {
                AssignAddress(value);
            }
        }

        /// <summary>
        /// Gets the link type.
        /// </summary>
        public TargetModeType LinkType
        {
            get
            {
                if (!string.IsNullOrEmpty(_model.SubAddress))
                {
                    return TargetModeType.CellReference;
                }

                var address = _model.Address;
                if (string.IsNullOrEmpty(address))
                {
                    return TargetModeType.External;
                }

                if (address.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                {
                    return TargetModeType.Email;
                }

                if (address.StartsWith("\\", StringComparison.Ordinal) || address.IndexOf(":\\", StringComparison.Ordinal) > 0)
                {
                    return TargetModeType.FilePath;
                }

                return TargetModeType.External;
            }
        }

        /// <summary>
        /// Gets or sets the screen tip.
        /// </summary>
        public string ScreenTip
        {
            get
            {
                return _model.ScreenTip ?? string.Empty;
            }
            set
            {
                _model.ScreenTip = NormalizeText(value);
            }
        }

        /// <summary>
        /// Gets or sets the text to display.
        /// </summary>
        public string TextToDisplay
        {
            get
            {
                return _model.TextToDisplay ?? string.Empty;
            }
            set
            {
                _model.TextToDisplay = NormalizeText(value);
            }
        }

        /// <summary>
        /// Performs delete.
        /// </summary>
        public void Delete()
        {
            _owner.Remove(_model);
        }

        private void AssignAddress(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                _model.Address = null;
                _model.SubAddress = null;
                return;
            }

            var normalized = value.Trim();
            if (normalized.StartsWith("#", StringComparison.Ordinal))
            {
                _model.Address = null;
                _model.SubAddress = normalized.Substring(1);
                return;
            }

            if (normalized.IndexOf('!') >= 0)
            {
                _model.Address = null;
                _model.SubAddress = normalized;
                return;
            }

            _model.Address = normalized;
            _model.SubAddress = null;
        }

        private static string NormalizeText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }

            return value;
        }
    }
}
