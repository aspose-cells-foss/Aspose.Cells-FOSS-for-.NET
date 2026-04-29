using System.IO;
using System.Collections.Generic;
using System;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Encapsulates a single worksheet and its supported v0.1 worksheet features.
    /// </summary>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// var sheet = workbook.Worksheets[0];
    ///
    /// sheet.Name = "Data";
    /// sheet.Cells["A1"].PutValue("North");
    /// sheet.Cells["B1"].PutValue(42);
    /// sheet.Zoom = 120;
    /// sheet.PageSetup.Orientation = PageOrientationType.Landscape;
    /// </code>
    /// </example>
    public class Worksheet
    {
        private readonly Workbook _workbook;
        private readonly WorksheetModel _model;
        private readonly Cells _cells;
        private readonly HyperlinkCollection _hyperlinks;
        private readonly ValidationCollection _validations;
        private readonly ConditionalFormattingCollection _conditionalFormattings;
        private readonly PageSetup _pageSetup;
        private readonly WorksheetProtection _protection;
        private readonly AutoFilter _autoFilter;
        private readonly ListObjectCollection _listObjects;
        private readonly PictureCollection _pictures;
        private readonly ShapeCollection _shapes;
        private readonly ChartCollection _charts;
        private readonly CommentCollection _comments;

        internal Worksheet(Workbook workbook, WorksheetModel model)
        {
            _workbook = workbook;
            _model = model;
            _cells = new Cells(this);
            _hyperlinks = new HyperlinkCollection(model.Hyperlinks);
            _validations = new ValidationCollection(model.Validations);
            _conditionalFormattings = new ConditionalFormattingCollection(model.ConditionalFormattings);
            _pageSetup = new PageSetup(model.PageSetup);
            _protection = new WorksheetProtection(model.Protection);
            _autoFilter = new AutoFilter(model.AutoFilter);
            _listObjects = new ListObjectCollection(model);
            _pictures = new PictureCollection(model);
            _shapes = new ShapeCollection(model);
            _charts = new ChartCollection(model);
            _comments = new CommentCollection(model.Comments);
        }

        internal WorksheetModel Model
        {
            get
            {
                return _model;
            }
        }

        internal Workbook Workbook
        {
            get
            {
                return _workbook;
            }
        }

        /// <summary>
        /// Gets or sets the worksheet name.
        /// </summary>
        public string Name
        {
            get
            {
                return _model.Name;
            }
            set
            {
                if (string.IsNullOrWhiteSpace(value)) throw new CellsException("Worksheet name must be non-empty.");
                _workbook.EnsureUniqueSheetName(value, _model);
                _model.Name = value;
            }
        }

        /// <summary>
        /// Gets or sets the worksheet visibility state.
        /// </summary>
        public VisibilityType VisibilityType
        {
            get
            {
                switch (_model.Visibility)
                {
                    case SheetVisibility.Hidden:
                        return Aspose.Cells_FOSS.VisibilityType.Hidden;
                    case SheetVisibility.VeryHidden:
                        return Aspose.Cells_FOSS.VisibilityType.VeryHidden;
                    default:
                        return Aspose.Cells_FOSS.VisibilityType.Visible;
                }
            }
            set
            {
                switch (value)
                {
                    case VisibilityType.Hidden:
                        _model.Visibility = SheetVisibility.Hidden;
                        break;
                    case VisibilityType.VeryHidden:
                        _model.Visibility = SheetVisibility.VeryHidden;
                        break;
                    default:
                        _model.Visibility = SheetVisibility.Visible;
                        break;
                }
            }
        }

        /// <summary>
        /// Gets or sets the worksheet tab color.
        /// </summary>
        public Color TabColor
        {
            get
            {
                if (_model.TabColor.HasValue)
                {
                    return Color.FromCore(_model.TabColor.Value);
                }

                return Color.Empty;
            }
            set
            {
                if (value.Equals(Color.Empty))
                {
                    _model.TabColor = null;
                    return;
                }

                _model.TabColor = value.ToCore();
            }
        }

        /// <summary>
        /// Gets or sets whether gridlines are shown in the worksheet view.
        /// </summary>
        public bool ShowGridlines
        {
            get
            {
                return _model.View.ShowGridLines;
            }
            set
            {
                _model.View.ShowGridLines = value;
            }
        }

        /// <summary>
        /// Gets or sets whether row and column headers are shown in the worksheet view.
        /// </summary>
        public bool ShowRowColumnHeaders
        {
            get
            {
                return _model.View.ShowRowColumnHeaders;
            }
            set
            {
                _model.View.ShowRowColumnHeaders = value;
            }
        }

        /// <summary>
        /// Gets or sets whether zero values are shown in the worksheet view.
        /// </summary>
        public bool ShowZeros
        {
            get
            {
                return _model.View.ShowZeros;
            }
            set
            {
                _model.View.ShowZeros = value;
            }
        }

        /// <summary>
        /// Gets or sets whether the worksheet view is right-to-left.
        /// </summary>
        public bool RightToLeft
        {
            get
            {
                return _model.View.RightToLeft;
            }
            set
            {
                _model.View.RightToLeft = value;
            }
        }

        /// <summary>
        /// Gets or sets the worksheet zoom percentage.
        /// </summary>
        public int Zoom
        {
            get
            {
                return _model.View.ZoomScale;
            }
            set
            {
                if (value < 10 || value > 400)
                {
                    throw new CellsException("Zoom must be between 10 and 400.");
                }

                _model.View.ZoomScale = value;
            }
        }

        /// <summary>
        /// Gets the cell grid facade for the worksheet.
        /// </summary>
        public Cells Cells
        {
            get
            {
                return _cells;
            }
        }

        /// <summary>
        /// Gets the worksheet hyperlink collection.
        /// </summary>
        public HyperlinkCollection Hyperlinks
        {
            get
            {
                return _hyperlinks;
            }
        }

        /// <summary>
        /// Gets the worksheet data validation collection.
        /// </summary>
        public ValidationCollection Validations
        {
            get
            {
                return _validations;
            }
        }

        /// <summary>
        /// Gets the worksheet conditional formatting collection.
        /// </summary>
        public ConditionalFormattingCollection ConditionalFormattings
        {
            get
            {
                return _conditionalFormattings;
            }
        }

        /// <summary>
        /// Gets page setup settings for the worksheet.
        /// </summary>
        public PageSetup PageSetup
        {
            get
            {
                return _pageSetup;
            }
        }

        /// <summary>
        /// Gets worksheet protection settings.
        /// </summary>
        public WorksheetProtection Protection
        {
            get
            {
                return _protection;
            }
        }

        /// <summary>
        /// Gets auto-filter settings for the worksheet.
        /// </summary>
        public AutoFilter AutoFilter
        {
            get
            {
                return _autoFilter;
            }
        }

        /// <summary>
        /// Gets the collection of Excel tables on this worksheet.
        /// </summary>
        public ListObjectCollection ListObjects
        {
            get
            {
                return _listObjects;
            }
        }

        /// <summary>
        /// Gets the collection of pictures on this worksheet.
        /// </summary>
        public PictureCollection Pictures
        {
            get
            {
                return _pictures;
            }
        }

        /// <summary>
        /// Gets the collection of drawing objects (shapes) on this worksheet.
        /// </summary>
        public ShapeCollection Shapes
        {
            get
            {
                return _shapes;
            }
        }

        /// <summary>
        /// Gets the collection of charts on this worksheet.
        /// </summary>
        public ChartCollection Charts
        {
            get
            {
                return _charts;
            }
        }

        /// <summary>
        /// Gets the collection of comments (legacy notes) on this worksheet.
        /// </summary>
        public CommentCollection Comments
        {
            get
            {
                return _comments;
            }
        }

        /// <summary>
        /// Marks the worksheet as protected using the current protection settings.
        /// </summary>
        public void Protect()
        {
            _model.Protection.IsProtected = true;
        }

        /// <summary>
        /// Clears worksheet protection and resets supported protection flags.
        /// </summary>
        public void Unprotect()
        {
            _protection.Reset();
        }
    }
}
