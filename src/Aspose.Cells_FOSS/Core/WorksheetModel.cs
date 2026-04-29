using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents worksheet model.
    /// </summary>
    public sealed class WorksheetModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetModel"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        public WorksheetModel(string name)
        {
            Name = name;
            Cells = new Dictionary<CellAddress, CellRecord>();
            Rows = new Dictionary<int, RowModel>();
            Columns = new List<ColumnRangeModel>();
            MergeRegions = new List<MergeRegion>();
            Hyperlinks = new List<HyperlinkModel>();
            Validations = new List<ValidationModel>();
            ConditionalFormattings = new List<ConditionalFormattingModel>();
            PageSetup = new PageSetupModel();
            View = new WorksheetViewModel();
            Protection = new WorksheetProtectionModel();
            AutoFilter = new AutoFilterModel();
            ListObjects = new List<ListObjectModel>();
            Pictures = new List<PictureModel>();
            ShapeImages = new List<ShapeImageModel>();
            Shapes = new List<ShapeModel>();
            Charts = new List<ChartModel>();
            Comments = new List<CommentModel>();
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the visibility.
        /// </summary>
        public SheetVisibility Visibility { get; set; } = SheetVisibility.Visible;
        /// <summary>
        /// Gets the cells.
        /// </summary>
        public Dictionary<CellAddress, CellRecord> Cells { get; }
        /// <summary>
        /// Gets the rows.
        /// </summary>
        public Dictionary<int, RowModel> Rows { get; }
        /// <summary>
        /// Gets the columns.
        /// </summary>
        public List<ColumnRangeModel> Columns { get; }
        /// <summary>
        /// Gets the merge regions.
        /// </summary>
        public List<MergeRegion> MergeRegions { get; }
        /// <summary>
        /// Gets the hyperlinks.
        /// </summary>
        public List<HyperlinkModel> Hyperlinks { get; }
        internal List<ValidationModel> Validations { get; }
        internal List<ConditionalFormattingModel> ConditionalFormattings { get; }
        /// <summary>
        /// Gets the page setup.
        /// </summary>
        public PageSetupModel PageSetup { get; }
        /// <summary>
        /// Gets the view.
        /// </summary>
        public WorksheetViewModel View { get; }
        /// <summary>
        /// Gets the protection.
        /// </summary>
        public WorksheetProtectionModel Protection { get; }
        /// <summary>
        /// Gets or sets the auto filter.
        /// </summary>
        public AutoFilterModel AutoFilter { get; }
        /// <summary>
        /// Gets the list of table models on this worksheet.
        /// </summary>
        public List<ListObjectModel> ListObjects { get; }
        /// <summary>
        /// Gets the list of picture models on this worksheet.
        /// </summary>
        public List<PictureModel> Pictures { get; }

        /// <summary>
        /// Gets the list of images embedded inside shape/group-shape raw XML (not top-level worksheet pictures).
        /// These are written to the drawing rels between pictures and charts so their original rId references remain valid.
        /// </summary>
        internal List<ShapeImageModel> ShapeImages { get; }
        /// <summary>
        /// Gets the list of shape models on this worksheet.
        /// </summary>
        public List<ShapeModel> Shapes { get; }
        /// <summary>
        /// Gets the list of chart models on this worksheet.
        /// </summary>
        public List<ChartModel> Charts { get; }
        /// <summary>
        /// Gets the list of comment models on this worksheet.
        /// </summary>
        public List<CommentModel> Comments { get; }
        /// <summary>
        /// Gets or sets the tab color.
        /// </summary>
        public ColorValue? TabColor { get; set; }
    }
}
