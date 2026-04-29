namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Specifies the chart type.
    /// </summary>
    public enum ChartType
    {
        // Standard 2D types
        /// <summary>Horizontal bar chart.</summary>
        Bar,
        /// <summary>Vertical column chart.</summary>
        Column,
        /// <summary>Line chart.</summary>
        Line,
        /// <summary>Area chart.</summary>
        Area,
        /// <summary>Pie chart.</summary>
        Pie,
        /// <summary>Doughnut chart.</summary>
        Doughnut,
        /// <summary>Scatter (XY) chart.</summary>
        Scatter,
        /// <summary>Bubble chart.</summary>
        Bubble,
        /// <summary>Radar/spider chart.</summary>
        Radar,
        /// <summary>Stock (OHLC) chart.</summary>
        Stock,
        // 3D types
        /// <summary>3D horizontal bar chart.</summary>
        Bar3D,
        /// <summary>3D vertical column chart.</summary>
        Column3D,
        /// <summary>3D line chart.</summary>
        Line3D,
        /// <summary>3D area chart.</summary>
        Area3D,
        /// <summary>3D pie chart.</summary>
        Pie3D,
        /// <summary>3D surface chart.</summary>
        Surface3D,
        /// <summary>3D wireframe surface chart.</summary>
        SurfaceWireframe3D,
        /// <summary>Contour (2D surface) chart.</summary>
        Contour,
        // Modern chartex types — loaded via raw XML only; Add() is not supported
        /// <summary>Waterfall chart (chartex).</summary>
        Waterfall,
        /// <summary>Treemap chart (chartex).</summary>
        Treemap,
        /// <summary>Sunburst chart (chartex).</summary>
        Sunburst,
        /// <summary>Histogram chart (chartex).</summary>
        Histogram,
        /// <summary>Box and whisker chart (chartex).</summary>
        BoxAndWhisker,
        /// <summary>Funnel chart (chartex).</summary>
        Funnel,
        /// <summary>Map chart (chartex).</summary>
        Map,
        /// <summary>Unrecognised or unsupported chart type.</summary>
        Unknown,
    }
}
