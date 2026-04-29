using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents a drawing object (auto shape) anchored to a worksheet.
    /// </summary>
    public sealed class Shape
    {
        private readonly ShapeModel _model;

        internal Shape(ShapeModel model)
        {
            _model = model;
        }

        /// <summary>
        /// Gets or sets the display name of the shape.
        /// </summary>
        public string Name
        {
            get
            {
                return _model.Name;
            }
            set
            {
                _model.Name = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based row index of the upper-left anchor cell.
        /// </summary>
        public int UpperLeftRow
        {
            get
            {
                return _model.UpperLeftRow;
            }
            set
            {
                if (value < 0)
                {
                    throw new CellsException("UpperLeftRow must be non-negative.");
                }

                _model.UpperLeftRow = value;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based column index of the upper-left anchor cell.
        /// </summary>
        public int UpperLeftColumn
        {
            get
            {
                return _model.UpperLeftColumn;
            }
            set
            {
                if (value < 0)
                {
                    throw new CellsException("UpperLeftColumn must be non-negative.");
                }

                _model.UpperLeftColumn = value;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based row index of the lower-right anchor cell.
        /// </summary>
        public int LowerRightRow
        {
            get
            {
                return _model.LowerRightRow;
            }
            set
            {
                _model.LowerRightRow = value;
            }
        }

        /// <summary>
        /// Gets or sets the zero-based column index of the lower-right anchor cell.
        /// </summary>
        public int LowerRightColumn
        {
            get
            {
                return _model.LowerRightColumn;
            }
            set
            {
                _model.LowerRightColumn = value;
            }
        }

        /// <summary>
        /// Gets or sets the preset geometry type as a raw DrawingML prst string (e.g. "rect", "rightArrow").
        /// Setting this also updates <see cref="AutoShapeType"/>.
        /// </summary>
        public string GeometryType
        {
            get
            {
                return _model.GeometryType;
            }
            set
            {
                _model.GeometryType = string.IsNullOrEmpty(value) ? "rect" : value;
            }
        }

        /// <summary>
        /// Gets or sets the shape type. Setting this updates <see cref="GeometryType"/>.
        /// </summary>
        public AutoShapeType AutoShapeType
        {
            get
            {
                return GeometryToAutoShapeType(_model.GeometryType);
            }
            set
            {
                _model.GeometryType = AutoShapeTypeToGeometry(value);
            }
        }

        internal static string AutoShapeTypeToGeometry(AutoShapeType type)
        {
            switch (type)
            {
                case AutoShapeType.Rectangle: return "rect";
                case AutoShapeType.RoundedRectangle: return "roundRect";
                case AutoShapeType.Ellipse: return "ellipse";
                case AutoShapeType.Triangle: return "triangle";
                case AutoShapeType.RightTriangle: return "rtTriangle";
                case AutoShapeType.Diamond: return "diamond";
                case AutoShapeType.Pentagon: return "pentagon";
                case AutoShapeType.Hexagon: return "hexagon";
                case AutoShapeType.Octagon: return "octagon";
                case AutoShapeType.Plus: return "plus";
                case AutoShapeType.Cube: return "cube";
                case AutoShapeType.Cylinder: return "cyl";
                case AutoShapeType.Heart: return "heart";
                case AutoShapeType.Lightning: return "lightningBolt";
                case AutoShapeType.Sun: return "sun";
                case AutoShapeType.Moon: return "moon";
                case AutoShapeType.Cloud: return "cloud";
                case AutoShapeType.RightArrow: return "rightArrow";
                case AutoShapeType.LeftArrow: return "leftArrow";
                case AutoShapeType.UpArrow: return "upArrow";
                case AutoShapeType.DownArrow: return "downArrow";
                case AutoShapeType.LeftRightArrow: return "leftRightArrow";
                case AutoShapeType.UpDownArrow: return "upDownArrow";
                case AutoShapeType.Star4Point: return "star4";
                case AutoShapeType.Star5Point: return "star5";
                case AutoShapeType.Star6Point: return "star6";
                case AutoShapeType.Star7Point: return "star7";
                case AutoShapeType.Star8Point: return "star8";
                case AutoShapeType.Star10Point: return "star10";
                case AutoShapeType.Star12Point: return "star12";
                case AutoShapeType.Star16Point: return "star16";
                case AutoShapeType.Star24Point: return "star24";
                case AutoShapeType.Star32Point: return "star32";
                case AutoShapeType.TextBox: return "rect";
                case AutoShapeType.MathPlus: return "mathPlus";
                case AutoShapeType.StraightConnector: return "straightConnector1";
                case AutoShapeType.BentConnector: return "bentConnector3";
                case AutoShapeType.CurvedConnector: return "curvedConnector3";
                default: return "rect";
            }
        }

        internal static AutoShapeType GeometryToAutoShapeType(string prst)
        {
            if (string.IsNullOrEmpty(prst))
            {
                return AutoShapeType.Unknown;
            }

            switch (prst)
            {
                case "rect": return AutoShapeType.Rectangle;
                case "roundRect": return AutoShapeType.RoundedRectangle;
                case "ellipse": return AutoShapeType.Ellipse;
                case "triangle": return AutoShapeType.Triangle;
                case "rtTriangle": return AutoShapeType.RightTriangle;
                case "diamond": return AutoShapeType.Diamond;
                case "pentagon": return AutoShapeType.Pentagon;
                case "hexagon": return AutoShapeType.Hexagon;
                case "octagon": return AutoShapeType.Octagon;
                case "plus": return AutoShapeType.Plus;
                case "cube": return AutoShapeType.Cube;
                case "cyl": return AutoShapeType.Cylinder;
                case "heart": return AutoShapeType.Heart;
                case "lightningBolt": return AutoShapeType.Lightning;
                case "sun": return AutoShapeType.Sun;
                case "moon": return AutoShapeType.Moon;
                case "cloud": return AutoShapeType.Cloud;
                case "rightArrow": return AutoShapeType.RightArrow;
                case "leftArrow": return AutoShapeType.LeftArrow;
                case "upArrow": return AutoShapeType.UpArrow;
                case "downArrow": return AutoShapeType.DownArrow;
                case "leftRightArrow": return AutoShapeType.LeftRightArrow;
                case "upDownArrow": return AutoShapeType.UpDownArrow;
                case "star4": return AutoShapeType.Star4Point;
                case "star5": return AutoShapeType.Star5Point;
                case "star6": return AutoShapeType.Star6Point;
                case "star7": return AutoShapeType.Star7Point;
                case "star8": return AutoShapeType.Star8Point;
                case "star10": return AutoShapeType.Star10Point;
                case "star12": return AutoShapeType.Star12Point;
                case "star16": return AutoShapeType.Star16Point;
                case "star24": return AutoShapeType.Star24Point;
                case "star32": return AutoShapeType.Star32Point;
                case "mathPlus": return AutoShapeType.MathPlus;
                case "straightConnector1": return AutoShapeType.StraightConnector;
                case "bentConnector2":
                case "bentConnector3":
                case "bentConnector4":
                case "bentConnector5": return AutoShapeType.BentConnector;
                case "curvedConnector2":
                case "curvedConnector3":
                case "curvedConnector4":
                case "curvedConnector5": return AutoShapeType.CurvedConnector;
                default: return AutoShapeType.Unknown;
            }
        }
    }
}
