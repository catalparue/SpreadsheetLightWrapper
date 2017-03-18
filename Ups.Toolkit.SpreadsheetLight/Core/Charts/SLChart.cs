using System.Collections.Generic;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using Ups.Toolkit.SpreadsheetLight.Core.style;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using Color = System.Drawing.Color;
using SLFill = Ups.Toolkit.SpreadsheetLight.Core.Drawing.SLFill;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     These correspond to the internal Open XML SDK classes
    /// </summary>
    internal enum SLInternalChartType
    {
        Area = 0,
        Area3D,
        Line,
        Line3D,
        Stock,
        Radar,
        Scatter,
        Pie,
        Pie3D,
        Doughnut,
        Bar,
        Bar3D,
        OfPie,
        Surface,
        Surface3D,
        Bubble
    }

    /// <summary>
    ///     Data series display types.
    /// </summary>
    public enum SLChartDataDisplayType
    {
        /// <summary>
        ///     Normal or clustered.
        /// </summary>
        Normal,

        /// <summary>
        ///     Stacked.
        /// </summary>
        Stacked,

        /// <summary>
        ///     100% stacked.
        /// </summary>
        StackedMax
    }

    /// <summary>
    ///     Built-in column chart types.
    /// </summary>
    public enum SLColumnChartType
    {
        /// <summary>
        ///     Clustered Column.
        /// </summary>
        ClusteredColumn = 0,

        /// <summary>
        ///     Stacked Column.
        /// </summary>
        StackedColumn,

        /// <summary>
        ///     100% Stacked Column.
        /// </summary>
        StackedColumnMax,

        /// <summary>
        ///     3D Clustered Column.
        /// </summary>
        ClusteredColumn3D,

        /// <summary>
        ///     Stacked Column in 3D.
        /// </summary>
        StackedColumn3D,

        /// <summary>
        ///     100% Stacked Column in 3D.
        /// </summary>
        StackedColumnMax3D,

        /// <summary>
        ///     3D Column.
        /// </summary>
        Column3D,

        /// <summary>
        ///     Clustered Cylinder.
        /// </summary>
        ClusteredCylinder,

        /// <summary>
        ///     Stacked Cylinder.
        /// </summary>
        StackedCylinder,

        /// <summary>
        ///     100% Stacked Cylinder.
        /// </summary>
        StackedCylinderMax,

        /// <summary>
        ///     3D Cylinder.
        /// </summary>
        Cylinder3D,

        /// <summary>
        ///     Clustered Cone.
        /// </summary>
        ClusteredCone,

        /// <summary>
        ///     Stacked Cone.
        /// </summary>
        StackedCone,

        /// <summary>
        ///     100% Stacked Cone.
        /// </summary>
        StackedConeMax,

        /// <summary>
        ///     3D Cone.
        /// </summary>
        Cone3D,

        /// <summary>
        ///     Clustered Pyramid.
        /// </summary>
        ClusteredPyramid,

        /// <summary>
        ///     Stacked Pyramid.
        /// </summary>
        StackedPyramid,

        /// <summary>
        ///     100% Stacked Pyramid.
        /// </summary>
        StackedPyramidMax,

        /// <summary>
        ///     3D Pyramid.
        /// </summary>
        Pyramid3D
    }

    /// <summary>
    ///     Built-in line chart types.
    /// </summary>
    public enum SLLineChartType
    {
        /// <summary>
        ///     Line.
        /// </summary>
        Line = 0,

        /// <summary>
        ///     Stacked Line.
        /// </summary>
        StackedLine,

        /// <summary>
        ///     100% Stacked Line.
        /// </summary>
        StackedLineMax,

        /// <summary>
        ///     Line with Markers.
        /// </summary>
        LineWithMarkers,

        /// <summary>
        ///     Stacked Line with Markers.
        /// </summary>
        StackedLineWithMarkers,

        /// <summary>
        ///     100% Stacked Line with Markers.
        /// </summary>
        StackedLineWithMarkersMax,

        /// <summary>
        ///     3D Line.
        /// </summary>
        Line3D
    }

    /// <summary>
    ///     Built-in pie chart types.
    /// </summary>
    public enum SLPieChartType
    {
        /// <summary>
        ///     Pie.
        /// </summary>
        Pie = 0,

        /// <summary>
        ///     Pie in 3D.
        /// </summary>
        Pie3D,

        /// <summary>
        ///     Pie of Pie.
        /// </summary>
        PieOfPie,

        /// <summary>
        ///     Exploded Pie.
        /// </summary>
        ExplodedPie,

        /// <summary>
        ///     Exploded Pie in 3D.
        /// </summary>
        ExplodedPie3D,

        /// <summary>
        ///     Bar of Pie
        /// </summary>
        BarOfPie
    }

    /// <summary>
    ///     Built-in bar chart types.
    /// </summary>
    public enum SLBarChartType
    {
        /// <summary>
        ///     Clustered Bar.
        /// </summary>
        ClusteredBar = 0,

        /// <summary>
        ///     Stacked Bar.
        /// </summary>
        StackedBar,

        /// <summary>
        ///     100% Stacked Bar.
        /// </summary>
        StackedBarMax,

        /// <summary>
        ///     Clustered Bar in 3D.
        /// </summary>
        ClusteredBar3D,

        /// <summary>
        ///     Stacked Bar in 3D.
        /// </summary>
        StackedBar3D,

        /// <summary>
        ///     100% Stacked Bar in 3D.
        /// </summary>
        StackedBarMax3D,

        /// <summary>
        ///     Clustered Horizontal Cylinder.
        /// </summary>
        ClusteredHorizontalCylinder,

        /// <summary>
        ///     Stacked Horizontal Cylinder.
        /// </summary>
        StackedHorizontalCylinder,

        /// <summary>
        ///     100% Stacked Horizontal Cylinder.
        /// </summary>
        StackedHorizontalCylinderMax,

        /// <summary>
        ///     Clustered Horizontal Cone.
        /// </summary>
        ClusteredHorizontalCone,

        /// <summary>
        ///     Stacked Horizontal Cone.
        /// </summary>
        StackedHorizontalCone,

        /// <summary>
        ///     100% Stacked Horizontal Cone.
        /// </summary>
        StackedHorizontalConeMax,

        /// <summary>
        ///     Clustered Horizontal Pyramid.
        /// </summary>
        ClusteredHorizontalPyramid,

        /// <summary>
        ///     Stacked Horizontal Pyramid.
        /// </summary>
        StackedHorizontalPyramid,

        /// <summary>
        ///     100% Stacked Horizontal Pyramid.
        /// </summary>
        StackedHorizontalPyramidMax
    }

    /// <summary>
    ///     Built-in area chart types.
    /// </summary>
    public enum SLAreaChartType
    {
        /// <summary>
        ///     Area.
        /// </summary>
        Area = 0,

        /// <summary>
        ///     Stacked Area.
        /// </summary>
        StackedArea,

        /// <summary>
        ///     100% Stacked Area.
        /// </summary>
        StackedAreaMax,

        /// <summary>
        ///     3D Area.
        /// </summary>
        Area3D,

        /// <summary>
        ///     Stacked Area in 3D.
        /// </summary>
        StackedArea3D,

        /// <summary>
        ///     100% Stacked Area in 3D.
        /// </summary>
        StackedAreaMax3D
    }

    /// <summary>
    ///     Built-in scatter chart types.
    /// </summary>
    public enum SLScatterChartType
    {
        /// <summary>
        ///     Scatter with only Markers.
        /// </summary>
        ScatterWithOnlyMarkers = 0,

        /// <summary>
        ///     Scatter with Smooth Lines and Markers.
        /// </summary>
        ScatterWithSmoothLinesAndMarkers,

        /// <summary>
        ///     Scatter with Smooth Lines.
        /// </summary>
        ScatterWithSmoothLines,

        /// <summary>
        ///     Scatter with Straight Lines and Markers.
        /// </summary>
        ScatterWithStraightLinesAndMarkers,

        /// <summary>
        ///     Scatter with Straight Lines.
        /// </summary>
        ScatterWithStraightLines
    }

    /// <summary>
    ///     Built-in stock chart types.
    /// </summary>
    public enum SLStockChartType
    {
        /// <summary>
        ///     High-Low-Close.
        /// </summary>
        HighLowClose = 0,

        /// <summary>
        ///     Open-High-Low-Close.
        /// </summary>
        OpenHighLowClose,

        /// <summary>
        ///     Volume-High-Low-Close.
        /// </summary>
        VolumeHighLowClose,

        /// <summary>
        ///     Volume-Open-High-Low-Close.
        /// </summary>
        VolumeOpenHighLowClose
    }

    /// <summary>
    ///     Built-in surface chart types.
    /// </summary>
    public enum SLSurfaceChartType
    {
        /// <summary>
        ///     3D Surface.
        /// </summary>
        Surface3D = 0,

        /// <summary>
        ///     Wiredframe 3D Surface.
        /// </summary>
        WireframeSurface3D,

        /// <summary>
        ///     Contour.
        /// </summary>
        Contour,

        /// <summary>
        ///     Wireframe Contour.
        /// </summary>
        WireframeContour
    }

    /// <summary>
    ///     Built-in doughnut chart types.
    /// </summary>
    public enum SLDoughnutChartType
    {
        /// <summary>
        ///     Doughnut.
        /// </summary>
        Doughnut = 0,

        /// <summary>
        ///     Exploded Doughnut.
        /// </summary>
        ExplodedDoughnut
    }

    /// <summary>
    ///     Built-in bubble chart types.
    /// </summary>
    public enum SLBubbleChartType
    {
        /// <summary>
        ///     Bubble.
        /// </summary>
        Bubble = 0,

        /// <summary>
        ///     Bubble with a 3D effect.
        /// </summary>
        Bubble3D
    }

    /// <summary>
    ///     Built-in radar chart types.
    /// </summary>
    public enum SLRadarChartType
    {
        /// <summary>
        ///     Radar.
        /// </summary>
        Radar = 0,

        /// <summary>
        ///     Radar with Markers.
        /// </summary>
        RadarWithMarkers,

        /// <summary>
        ///     Filled Radar.
        /// </summary>
        FilledRadar
    }

    /// <summary>
    ///     Built-in chart styles.
    /// </summary>
    public enum SLChartStyle : byte
    {
        // the numbers assigned have to be those assigned as follows.

        /// <summary>
        ///     Standard style in black and white.
        /// </summary>
        Style1 = 1,

        /// <summary>
        ///     Standard style in theme colors. This is the default.
        /// </summary>
        Style2 = 2,

        /// <summary>
        ///     Standard style in tints of accent 1 color.
        /// </summary>
        Style3 = 3,

        /// <summary>
        ///     Standard style in tints of accent 2 color.
        /// </summary>
        Style4 = 4,

        /// <summary>
        ///     Standard style in tints of accent 3 color.
        /// </summary>
        Style5 = 5,

        /// <summary>
        ///     Standard style in tints of accent 4 color.
        /// </summary>
        Style6 = 6,

        /// <summary>
        ///     Standard style in tints of accent 5 color.
        /// </summary>
        Style7 = 7,

        /// <summary>
        ///     Standard style in tints of accent 6 color.
        /// </summary>
        Style8 = 8,

        /// <summary>
        ///     Bordered data series in black and white.
        /// </summary>
        Style9 = 9,

        /// <summary>
        ///     Bordered data series in theme colors.
        /// </summary>
        Style10 = 10,

        /// <summary>
        ///     Bordered data series in tints of accent 1 color.
        /// </summary>
        Style11 = 11,

        /// <summary>
        ///     Bordered data series in tints of accent 2 color.
        /// </summary>
        Style12 = 12,

        /// <summary>
        ///     Bordered data series in tints of accent 3 color.
        /// </summary>
        Style13 = 13,

        /// <summary>
        ///     Bordered data series in tints of accent 4 color.
        /// </summary>
        Style14 = 14,

        /// <summary>
        ///     Bordered data series in tints of accent 5 color.
        /// </summary>
        Style15 = 15,

        /// <summary>
        ///     Bordered data series in tints of accent 6 color.
        /// </summary>
        Style16 = 16,

        /// <summary>
        ///     Softly blurred data series in black and white.
        /// </summary>
        Style17 = 17,

        /// <summary>
        ///     Softly blurred data series in theme colors.
        /// </summary>
        Style18 = 18,

        /// <summary>
        ///     Softly blurred data series in tints of accent 1 color.
        /// </summary>
        Style19 = 19,

        /// <summary>
        ///     Softly blurred data series in tints of accent 2 color.
        /// </summary>
        Style20 = 20,

        /// <summary>
        ///     Softly blurred data series in tints of accent 3 color.
        /// </summary>
        Style21 = 21,

        /// <summary>
        ///     Softly blurred data series in tints of accent 4 color.
        /// </summary>
        Style22 = 22,

        /// <summary>
        ///     Softly blurred data series in tints of accent 5 color.
        /// </summary>
        Style23 = 23,

        /// <summary>
        ///     Softly blurred data series in tints of accent 6 color.
        /// </summary>
        Style24 = 24,

        /// <summary>
        ///     Bevelled data series in black and white.
        /// </summary>
        Style25 = 25,

        /// <summary>
        ///     Bevelled data series in theme colors.
        /// </summary>
        Style26 = 26,

        /// <summary>
        ///     Bevelled data series in tints of accent 1 color.
        /// </summary>
        Style27 = 27,

        /// <summary>
        ///     Bevelled data series in tints of accent 2 color.
        /// </summary>
        Style28 = 28,

        /// <summary>
        ///     Bevelled data series in tints of accent 3 color.
        /// </summary>
        Style29 = 29,

        /// <summary>
        ///     Bevelled data series in tints of accent 4 color.
        /// </summary>
        Style30 = 30,

        /// <summary>
        ///     Bevelled data series in tints of accent 5 color.
        /// </summary>
        Style31 = 31,

        /// <summary>
        ///     Bevelled data series in tints of accent 6 color.
        /// </summary>
        Style32 = 32,

        /// <summary>
        ///     Standard style in black and white, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style33 = 33,

        /// <summary>
        ///     Standard style in theme colors, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style34 = 34,

        /// <summary>
        ///     Standard style in tints of accent 1 color, with gray-filled plot area (side wall, back wall and floor for 3D
        ///     charts).
        /// </summary>
        Style35 = 35,

        /// <summary>
        ///     Standard style in tints of accent 2 color, with gray-filled plot area (side wall, back wall and floor for 3D
        ///     charts).
        /// </summary>
        Style36 = 36,

        /// <summary>
        ///     Standard style in tints of accent 3 color, with gray-filled plot area (side wall, back wall and floor for 3D
        ///     charts).
        /// </summary>
        Style37 = 37,

        /// <summary>
        ///     Standard style in tints of accent 4 color, with gray-filled plot area (side wall, back wall and floor for 3D
        ///     charts).
        /// </summary>
        Style38 = 38,

        /// <summary>
        ///     Standard style in tints of accent 5 color, with gray-filled plot area (side wall, back wall and floor for 3D
        ///     charts).
        /// </summary>
        Style39 = 39,

        /// <summary>
        ///     Standard style in tints of accent 6 color, with gray-filled plot area (side wall, back wall and floor for 3D
        ///     charts).
        /// </summary>
        Style40 = 40,

        /// <summary>
        ///     Softly blurred and bevelled data series in black and white, with black chart area and gray-filled plot area (side
        ///     wall, back wall and floor for 3D charts).
        /// </summary>
        Style41 = 41,

        /// <summary>
        ///     Softly blurred and bevelled data series in theme colors, with black chart area and gray-filled plot area (side
        ///     wall, back wall and floor for 3D charts).
        /// </summary>
        Style42 = 42,

        /// <summary>
        ///     Softly blurred and bevelled data series in tints of accent 1 color, with black chart area and gray-filled plot area
        ///     (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style43 = 43,

        /// <summary>
        ///     Softly blurred and bevelled data series in tints of accent 2 color, with black chart area and gray-filled plot area
        ///     (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style44 = 44,

        /// <summary>
        ///     Softly blurred and bevelled data series in tints of accent 3 color, with black chart area and gray-filled plot area
        ///     (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style45 = 45,

        /// <summary>
        ///     Softly blurred and bevelled data series in tints of accent 4 color, with black chart area and gray-filled plot area
        ///     (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style46 = 46,

        /// <summary>
        ///     Softly blurred and bevelled data series in tints of accent 5 color, with black chart area and gray-filled plot area
        ///     (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style47 = 47,

        /// <summary>
        ///     Softly blurred and bevelled data series in tints of accent 6 color, with black chart area and gray-filled plot area
        ///     (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style48 = 48
    }

    // this is for ChartSpace root class

    /// <summary>
    ///     Encapsulates properties and methods for a chart to be inserted into a worksheet.
    /// </summary>
    public class SLChart
    {
        // used to moderate the behaviour of the secondary text axis.
        // Initially, the axis is at the bottom (if not bar chart), but deleted.
        // Then when shown, the axis goes to the top. If then hidden, the axis stays at the top but is deleted.
        internal bool HasShownSecondaryTextAxis;

        internal bool Is3D;
        internal List<Color> listThemeColors;

        internal SLShapeProperties ShapeProperties;

        internal SLChart()
        {
        }

        internal bool Date1904 { get; set; }

        /// <summary>
        ///     True if follow latest Excel styling defaults (but no guarantees because I might not
        ///     be able to afford to keep buying latest Office/Excel).
        /// </summary>
        internal bool IsStylish { get; set; }

        /// <summary>
        ///     Specifies whether the chart has rounded corners. In Microsoft Excel, you might find this setting under "Border
        ///     Styles" when formatting the chart area.
        /// </summary>
        public bool RoundedCorners { get; set; }

        internal bool IsCombinable { get; set; }

        internal double TopPosition { get; set; }
        internal double LeftPosition { get; set; }
        internal double BottomPosition { get; set; }
        internal double RightPosition { get; set; }

        // this is the primary data source
        internal string WorksheetName { get; set; }
        internal bool RowsAsDataSeries { get; set; }
        internal bool ShowHiddenData { get; set; }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal SLChartStyle ChartStyle { get; set; }

        /// <summary>
        ///     The default is to show empty cells with a gap (or whichever option is appropriate for the chart). Note that "Zero"
        ///     and "Span" are used mostly for line, scatter and radar charts. Use "Zero" to force a zero value, and "Span" to
        ///     connect data points across the empty cell.
        /// </summary>
        public C.DisplayBlanksAsValues ShowEmptyCellsAs { get; set; }

        /// <summary>
        ///     Indicates whether data labels over the maximum value of the chart is shown. The default value is true.
        /// </summary>
        public bool ShowDataLabelsOverMaximum { get; set; }

        internal bool HasView3D
        {
            get
            {
                return (RotateX != null) || (HeightPercent != null) || (RotateY != null) || (DepthPercent != null) ||
                       (RightAngleAxes != null) || (Perspective != null);
            }
        }

        // RotateX and RotateY don't correspond to the X- and Y-axis on the Excel user interface.
        // Why? Why?!? WHY?!?! I don't know. Go ask Microsoft...
        internal sbyte? RotateX { get; set; }
        internal ushort? HeightPercent { get; set; }
        internal ushort? RotateY { get; set; }
        internal ushort? DepthPercent { get; set; }
        internal bool? RightAngleAxes { get; set; }

        /// <summary>
        ///     This is double that's shown in Excel. Excel values range from 0 to 120 degrees.
        ///     So this is 0 to 240 units. "Default" rotation angle is 30 (15 degrees).
        ///     Did Microsoft want to make full use of the byte range value?
        /// </summary>
        internal byte? Perspective { get; set; }

        /// <summary>
        ///     A friendly name for the chart. By default, this is in the form of "Chart #", where "#" is a number.
        /// </summary>
        public string ChartName { get; set; }

        internal bool HasTitle { get; set; }

        /// <summary>
        ///     The chart title. By default the chart title is hidden, so make sure to show it if chart title properties are set.
        /// </summary>
        public SLTitle Title { get; set; }

        /// <summary>
        ///     The floor of 3D charts.
        /// </summary>
        public SLFloor Floor { get; set; }

        /// <summary>
        ///     The side wall of 3D charts. Note that contour charts don't show the side wall, even though they're technically 3D
        ///     charts.
        /// </summary>
        public SLSideWall SideWall { get; set; }

        /// <summary>
        ///     The back wall of 3D charts. Note that contour charts don't show the back wall, even though they're technically 3D
        ///     charts.
        /// </summary>
        public SLBackWall BackWall { get; set; }

        /// <summary>
        ///     The plot area.
        /// </summary>
        public SLPlotArea PlotArea { get; set; }

        /// <summary>
        ///     The primary chart text axis. This is usually the horizontal axis at the bottom (bar charts have them on the left).
        ///     Depending on the type of chart, this can be a category, date or value axis.
        /// </summary>
        public SLTextAxis PrimaryTextAxis
        {
            get { return PlotArea.PrimaryTextAxis; }
        }

        /// <summary>
        ///     The primary chart value axis. This is usually the vertical axis on the left (bar charts have them at the bottom).
        /// </summary>
        public SLValueAxis PrimaryValueAxis
        {
            get { return PlotArea.PrimaryValueAxis; }
        }

        /// <summary>
        ///     The depth axis for 3D charts.
        /// </summary>
        public SLSeriesAxis DepthAxis
        {
            get { return PlotArea.DepthAxis; }
        }

        /// <summary>
        ///     The secondary chart text axis. This is usually the horizontal axis at the top (bar charts have them on the left
        ///     initially until you show this axis).
        ///     Depending on the type of chart, this can be a category, date or value axis.
        /// </summary>
        public SLTextAxis SecondaryTextAxis
        {
            get { return PlotArea.SecondaryTextAxis; }
        }

        /// <summary>
        ///     The secondary chart value axis. This is usually the vertical axis on the right (bar charts have them at the top).
        /// </summary>
        public SLValueAxis SecondaryValueAxis
        {
            get { return PlotArea.SecondaryValueAxis; }
        }

        /// <summary>
        ///     Specifies if the data table is shown.
        /// </summary>
        public bool ShowDataTable
        {
            get { return PlotArea.ShowDataTable; }
            set { PlotArea.ShowDataTable = value; }
        }

        /// <summary>
        ///     The data table of the chart.
        /// </summary>
        public SLDataTable DataTable
        {
            get { return PlotArea.DataTable; }
        }

        internal bool ShowLegend { get; set; }

        /// <summary>
        ///     The chart legend.
        /// </summary>
        public SLLegend Legend { get; set; }

        /// <summary>
        ///     Fill properties.
        /// </summary>
        public SLFill Fill
        {
            get { return ShapeProperties.Fill; }
        }

        /// <summary>
        ///     Border properties.
        /// </summary>
        public SLLinePropertiesType Border
        {
            get { return ShapeProperties.Outline; }
        }

        /// <summary>
        ///     Shadow properties.
        /// </summary>
        public SLShadowEffect Shadow
        {
            get { return ShapeProperties.EffectList.Shadow; }
        }

        /// <summary>
        ///     Glow properties.
        /// </summary>
        public SLGlow Glow
        {
            get { return ShapeProperties.EffectList.Glow; }
        }

        /// <summary>
        ///     Soft edge properties.
        /// </summary>
        public SLSoftEdge SoftEdge
        {
            get { return ShapeProperties.EffectList.SoftEdge; }
        }

        /// <summary>
        ///     3D format properties.
        /// </summary>
        public SLFormat3D Format3D
        {
            get { return ShapeProperties.Format3D; }
        }

        /// <summary>
        ///     Set the chart style using one of the built-in styles. WARNING: This is supposedly phased out in Excel 2013. Maybe
        ///     it'll be replaced by something else, maybe not at all.
        /// </summary>
        /// <param name="ChartStyle">A built-in chart style.</param>
        public void SetChartStyle(SLChartStyle ChartStyle)
        {
            this.ChartStyle = ChartStyle;
        }

        /// <summary>
        ///     Set a pie chart using one of the built-in pie chart types.
        /// </summary>
        /// <param name="ChartType">A built-in pie chart type.</param>
        public void SetChartType(SLPieChartType ChartType)
        {
            SetChartType(ChartType, null);
        }

        /// <summary>
        ///     Set a pie chart using one of the built-in pie chart types.
        /// </summary>
        /// <param name="ChartType">A built-in pie chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLPieChartType ChartType, SLPieChartOptions Options)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLPieChartType.Pie:
                    vType = SLDataSeriesChartType.PieChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);
                    break;
                case SLPieChartType.Pie3D:
                    RotateX = 30;
                    if (Options != null)
                        RotateY = Options.FirstSliceAngle;
                    else
                        RotateY = 0;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Pie3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);
                    break;
                case SLPieChartType.PieOfPie:
                    vType = SLDataSeriesChartType.OfPieChartPie;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    PlotArea.UsedChartOptions[iChartType].GapWidth = 100;
                    PlotArea.UsedChartOptions[iChartType].SecondPieSize = 75;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);
                    break;
                case SLPieChartType.ExplodedPie:
                    vType = SLDataSeriesChartType.PieChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Explosion = 25;
                    break;
                case SLPieChartType.ExplodedPie3D:
                    RotateX = 30;
                    if (Options != null)
                        RotateY = Options.FirstSliceAngle;
                    else
                        RotateY = 0;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Pie3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Explosion = 25;
                    break;
                case SLPieChartType.BarOfPie:
                    vType = SLDataSeriesChartType.OfPieChartBar;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    PlotArea.UsedChartOptions[iChartType].GapWidth = 100;
                    PlotArea.UsedChartOptions[iChartType].SecondPieSize = 75;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);
                    break;
            }
        }

        /// <summary>
        ///     Set a surface chart using one of the built-in surface chart types.
        /// </summary>
        /// <param name="ChartType">A built-in surface chart type.</param>
        public void SetChartType(SLSurfaceChartType ChartType)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLSurfaceChartType.Surface3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Surface3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
                    PlotArea.HasDepthAxis = true;
                    PlotArea.DepthAxis.IsCrosses = true;
                    PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
                case SLSurfaceChartType.WireframeSurface3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Surface3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Wireframe = true;
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
                    PlotArea.HasDepthAxis = true;
                    PlotArea.DepthAxis.IsCrosses = true;
                    PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
                case SLSurfaceChartType.Contour:
                    RotateX = 90;
                    RotateY = 0;
                    RightAngleAxes = false;
                    Perspective = 0;

                    vType = SLDataSeriesChartType.SurfaceChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
                    PlotArea.HasDepthAxis = true;
                    PlotArea.DepthAxis.IsCrosses = true;
                    PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
                case SLSurfaceChartType.WireframeContour:
                    RotateX = 90;
                    RotateY = 0;
                    RightAngleAxes = false;
                    Perspective = 0;

                    vType = SLDataSeriesChartType.Surface3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Wireframe = true;
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
                    PlotArea.HasDepthAxis = true;
                    PlotArea.DepthAxis.IsCrosses = true;
                    PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
            }
        }

        /// <summary>
        ///     Set a radar chart using one of the built-in radar chart types.
        /// </summary>
        /// <param name="ChartType">A built-in radar chart type.</param>
        public void SetChartType(SLRadarChartType ChartType)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLRadarChartType.Radar:
                    vType = SLDataSeriesChartType.RadarChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Marker;
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;

                    if (IsStylish)
                        Legend.LegendPosition = C.LegendPositionValues.Top;

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;
                    break;
                case SLRadarChartType.RadarWithMarkers:
                    vType = SLDataSeriesChartType.RadarChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Marker;
                    PlotArea.SetDataSeriesChartType(vType);

                    if (IsStylish)
                    {
                        for (var i = 0; i < PlotArea.DataSeries.Count; ++i)
                        {
                            PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }
                        Legend.LegendPosition = C.LegendPositionValues.Top;
                    }

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;
                    break;
                case SLRadarChartType.FilledRadar:
                    vType = SLDataSeriesChartType.RadarChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Filled;
                    PlotArea.SetDataSeriesChartType(vType);

                    if (IsStylish)
                        Legend.LegendPosition = C.LegendPositionValues.Top;

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;
                    break;
            }
        }

        // TODO! What's the correct bubble chart behaviour?

        /// <summary>
        ///     Set a bubble chart using one of the built-in bubble chart types.
        /// </summary>
        /// <param name="ChartType">A built-in bubble chart type.</param>
        public void SetChartType(SLBubbleChartType ChartType)
        {
            SetChartType(ChartType, null);
        }

        /// <summary>
        ///     Set a bubble chart using one of the built-in bubble chart types.
        /// </summary>
        /// <param name="ChartType">A built-in bubble chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLBubbleChartType ChartType, SLBubbleChartOptions Options)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            int i, index;
            var nl = new SLNumberLiteral();
            nl.FormatCode = SLConstants.NumberFormatGeneral;
            nl.PointCount = (uint) (EndRowIndex - StartRowIndex);
            for (i = 0; i < EndRowIndex - StartRowIndex; ++i)
                nl.Points.Add(new SLNumericPoint {Index = (uint) i, NumericValue = "1"});

            double fTemp = 0;

            var series = new List<SLDataSeries>();
            SLDataSeries ser;
            for (index = 0, i = 0; i < PlotArea.DataSeries.Count; ++index, ++i)
            {
                ser = new SLDataSeries(listThemeColors);
                ser.Index = (uint) index;
                ser.Order = (uint) index;
                ser.IsStringReference = null;
                ser.StringReference = PlotArea.DataSeries[i].StringReference.Clone();

                ser.AxisData = PlotArea.DataSeries[i].AxisData.Clone();
                if (PlotArea.DataSeries[i].StringReference.Points.Count > 0)
                {
                    foreach (var pt in ser.AxisData.StringReference.Points)
                        ++pt.Index;

                    // move one row up
                    --ser.AxisData.StringReference.StartRowIndex;
                    ser.AxisData.StringReference.RefreshFormula();

                    ser.AxisData.StringReference.Points.Insert(0,
                        PlotArea.DataSeries[i].StringReference.Points[0].Clone());
                    ++ser.AxisData.StringReference.PointCount;
                }
                ser.NumberData = PlotArea.DataSeries[i].NumberData.Clone();
                if (PlotArea.DataSeries[i].StringReference.Points.Count > 1)
                {
                    foreach (var pt in ser.NumberData.NumberReference.NumberingCache.Points)
                        ++pt.Index;

                    --ser.NumberData.NumberReference.StartRowIndex;
                    ser.NumberData.NumberReference.RefreshFormula();

                    if (double.TryParse(PlotArea.DataSeries[i].StringReference.Points[1].NumericValue, NumberStyles.Any,
                        CultureInfo.InvariantCulture, out fTemp))
                        ser.NumberData.NumberReference.NumberingCache.Points.Insert(0,
                            new SLNumericPoint {Index = 0, NumericValue = fTemp.ToString(CultureInfo.InvariantCulture)});
                    else
                        ser.NumberData.NumberReference.NumberingCache.Points.Insert(0,
                            new SLNumericPoint {Index = 0, NumericValue = "0"});
                    ++ser.NumberData.NumberReference.NumberingCache.PointCount;
                }

                ++i;
                if (i < PlotArea.DataSeries.Count)
                {
                    ser.BubbleSize = PlotArea.DataSeries[i].NumberData.Clone();

                    if (PlotArea.DataSeries[i].StringReference.Points.Count > 2)
                    {
                        foreach (var pt in ser.BubbleSize.NumberReference.NumberingCache.Points)
                            ++pt.Index;

                        --ser.BubbleSize.NumberReference.StartRowIndex;
                        ser.BubbleSize.NumberReference.RefreshFormula();

                        if (double.TryParse(PlotArea.DataSeries[i].StringReference.Points[2].NumericValue,
                            NumberStyles.Any, CultureInfo.InvariantCulture, out fTemp))
                            ser.BubbleSize.NumberReference.NumberingCache.Points.Insert(0,
                                new SLNumericPoint
                                {
                                    Index = 0,
                                    NumericValue = fTemp.ToString(CultureInfo.InvariantCulture)
                                });
                        else
                            ser.BubbleSize.NumberReference.NumberingCache.Points.Insert(0,
                                new SLNumericPoint {Index = 0, NumericValue = "0"});
                        ++ser.BubbleSize.NumberReference.NumberingCache.PointCount;
                    }
                }
                else
                {
                    ser.BubbleSize.UseNumberLiteral = true;
                    ser.BubbleSize.NumberLiteral = nl.Clone();
                }
                series.Add(ser);
            }

            PlotArea.DataSeries = series;

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLBubbleChartType.Bubble:
                    vType = SLDataSeriesChartType.BubbleChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BubbleScale = 100;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Bubble3D = false;

                    SetPlotAreaValueAxes();
                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLBubbleChartType.Bubble3D:
                    vType = SLDataSeriesChartType.BubbleChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BubbleScale = 100;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Bubble3D = true;

                    SetPlotAreaValueAxes();
                    PlotArea.HasPrimaryAxes = true;
                    break;
            }
        }

        /// <summary>
        ///     Set a stock chart using one of the built-in stock chart types.
        /// </summary>
        /// <param name="ChartType">A built-in stock chart type.</param>
        public void SetChartType(SLStockChartType ChartType)
        {
            SetChartType(ChartType, null);
        }

        /// <summary>
        ///     Set a stock chart using one of the built-in stock chart types.
        /// </summary>
        /// <param name="ChartType">A built-in stock chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLStockChartType ChartType, SLStockChartOptions Options)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            int i;
            int iBarChartType;

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLStockChartType.HighLowClose:
                    vType = SLDataSeriesChartType.StockChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                    for (i = 0; i < PlotArea.DataSeries.Count; ++i)
                    {
                        PlotArea.DataSeries[i].ChartType = vType;
                        if (IsStylish)
                        {
                            PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                            PlotArea.DataSeries[i].Options.Line.CapType = A.LineCapValues.Round;
                            PlotArea.DataSeries[i].Options.Line.SetNoLine();
                            PlotArea.DataSeries[i].Options.Line.JoinType = SLLineJoinValues.Round;
                            PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        }
                    }

                    // this is for Close
                    if (PlotArea.DataSeries.Count > 2)
                    {
                        PlotArea.DataSeries[2].Options.Marker.Symbol = C.MarkerStyleValues.Dot;
                        PlotArea.DataSeries[2].Options.Marker.Size = 3;
                        if (IsStylish)
                        {
                            PlotArea.DataSeries[2].Options.Marker.Fill.SetSolidFill(A.SchemeColorValues.Accent3, 0, 0);
                            PlotArea.DataSeries[2].Options.Marker.Line.Width = 0.75m;
                            PlotArea.DataSeries[2].Options.Marker.Line.SetSolidLine(A.SchemeColorValues.Accent3, 0, 0);
                        }
                    }

                    PlotArea.UsedChartOptions[iChartType].HasHighLowLines = true;
                    if (IsStylish)
                    {
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Width = 0.75m;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CapType = A.LineCapValues.Flat;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CompoundLineType =
                            A.CompoundLineValues.Single;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Alignment = A.PenAlignmentValues.Center;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.SetSolidLine(A.SchemeColorValues.Text1,
                            0.25m, 0);
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.JoinType = SLLineJoinValues.Round;
                    }

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.ShowMajorGridlines = false;
                        PlotArea.PrimaryTextAxis.Fill.SetNoFill();
                        // 2.25 pt width
                        PlotArea.PrimaryTextAxis.Line.Width = 0.75m;
                        PlotArea.PrimaryTextAxis.Line.CapType = A.LineCapValues.Flat;
                        PlotArea.PrimaryTextAxis.Line.CompoundLineType = A.CompoundLineValues.Single;
                        PlotArea.PrimaryTextAxis.Line.Alignment = A.PenAlignmentValues.Center;
                        PlotArea.PrimaryTextAxis.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                        PlotArea.PrimaryTextAxis.Line.JoinType = SLLineJoinValues.Round;
                        var rst = new SLRstType();
                        rst.AppendText(" ", new SLFont
                        {
                            FontScheme = FontSchemeValues.Minor,
                            FontSize = 9,
                            Bold = false,
                            Italic = false,
                            Underline = UnderlineValues.None,
                            Strike = false
                        });
                        PlotArea.PrimaryTextAxis.Title.SetTitle(rst);
                        PlotArea.PrimaryTextAxis.Title.Fill.SetSolidFill(A.SchemeColorValues.Text1, 0.35m, 0);

                        PlotArea.PrimaryValueAxis.MinorTickMark = C.TickMarkValues.None;
                    }

                    if (IsStylish) Legend.LegendPosition = C.LegendPositionValues.Bottom;
                    PlotArea.SetDataSeriesAutoAxisType();
                    ShowEmptyCellsAs = C.DisplayBlanksAsValues.Gap;
                    break;
                case SLStockChartType.OpenHighLowClose:
                    vType = SLDataSeriesChartType.StockChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                    for (i = 0; i < PlotArea.DataSeries.Count; ++i)
                    {
                        PlotArea.DataSeries[i].ChartType = vType;
                        if (IsStylish)
                        {
                            // 2.25 pt width
                            PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                            PlotArea.DataSeries[i].Options.Line.CapType = A.LineCapValues.Round;
                            PlotArea.DataSeries[i].Options.Line.SetNoLine();
                            PlotArea.DataSeries[i].Options.Line.JoinType = SLLineJoinValues.Round;
                            PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        }
                    }

                    PlotArea.UsedChartOptions[iChartType].HasHighLowLines = true;
                    PlotArea.UsedChartOptions[iChartType].HasUpDownBars = true;
                    if (IsStylish)
                    {
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Width = 0.75m;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CapType = A.LineCapValues.Flat;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CompoundLineType =
                            A.CompoundLineValues.Single;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Alignment = A.PenAlignmentValues.Center;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.SetSolidLine(A.SchemeColorValues.Text1,
                            0.25m, 0);
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.JoinType = SLLineJoinValues.Round;

                        PlotArea.UsedChartOptions[iChartType].UpDownBars.GapWidth = 150;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Fill.SetSolidFill(
                            A.SchemeColorValues.Light1, 0, 0);
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.Width = 0.75m;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.CapType = A.LineCapValues.Flat;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.CompoundLineType =
                            A.CompoundLineValues.Single;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.Alignment =
                            A.PenAlignmentValues.Center;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.SetSolidLine(
                            A.SchemeColorValues.Text1, 0.35m, 0);
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.JoinType =
                            SLLineJoinValues.Round;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Fill.SetSolidFill(
                            A.SchemeColorValues.Dark1, 0.25m, 0);
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.Width = 0.75m;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.CapType = A.LineCapValues.Flat;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.CompoundLineType =
                            A.CompoundLineValues.Single;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.Alignment =
                            A.PenAlignmentValues.Center;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.SetSolidLine(
                            A.SchemeColorValues.Text1, 0.35m, 0);
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.JoinType =
                            SLLineJoinValues.Round;
                    }

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;

                    if (IsStylish) Legend.LegendPosition = C.LegendPositionValues.Bottom;
                    PlotArea.SetDataSeriesAutoAxisType();
                    ShowEmptyCellsAs = C.DisplayBlanksAsValues.Gap;
                    break;
                case SLStockChartType.VolumeHighLowClose:
                    vType = SLDataSeriesChartType.StockChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                    iBarChartType = (int) SLDataSeriesChartType.BarChartColumnPrimary;
                    for (i = 0; i < PlotArea.DataSeries.Count; ++i)
                        if (i == 0)
                        {
                            PlotArea.DataSeries[i].ChartType = SLDataSeriesChartType.BarChartColumnPrimary;
                            if (IsStylish)
                            {
                                PlotArea.DataSeries[i].Options.Fill.SetSolidFill(A.SchemeColorValues.Accent1, 0, 0);
                                // 2.25 pt width
                                PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                                PlotArea.DataSeries[i].Options.Line.SetNoLine();
                            }

                            PlotArea.UsedChartTypes[iBarChartType] = true;
                            PlotArea.UsedChartOptions[iBarChartType].BarDirection = C.BarDirectionValues.Column;
                            PlotArea.UsedChartOptions[iBarChartType].BarGrouping = C.BarGroupingValues.Clustered;
                            if (Options != null)
                            {
                                PlotArea.UsedChartOptions[iBarChartType].GapWidth = Options.GapWidth;
                                PlotArea.UsedChartOptions[iBarChartType].Overlap = Options.Overlap;
                                PlotArea.DataSeries[i].Options.ShapeProperties.Fill = Options.Fill.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.Outline = Options.Border.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.Shadow =
                                    Options.Shadow.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.Glow = Options.Glow.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.SoftEdge =
                                    Options.SoftEdge.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.Format3D = Options.Format3D.Clone();
                            }
                        }
                        else
                        {
                            PlotArea.DataSeries[i].ChartType = vType;
                            if (IsStylish)
                            {
                                // 2.25 pt width
                                PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                                PlotArea.DataSeries[i].Options.Line.CapType = A.LineCapValues.Round;
                                PlotArea.DataSeries[i].Options.Line.SetNoLine();
                                PlotArea.DataSeries[i].Options.Line.JoinType = SLLineJoinValues.Round;
                                PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.None;
                            }
                        }

                    // this is for Close
                    if (PlotArea.DataSeries.Count > 3)
                    {
                        PlotArea.DataSeries[3].Options.Marker.Symbol = C.MarkerStyleValues.Dot;
                        PlotArea.DataSeries[3].Options.Marker.Size = 5;
                        if (IsStylish)
                        {
                            PlotArea.DataSeries[3].Options.Marker.Fill.SetSolidFill(A.SchemeColorValues.Accent4, 0, 0);
                            PlotArea.DataSeries[3].Options.Marker.Line.Width = 0.75m;
                            PlotArea.DataSeries[3].Options.Marker.Line.SetSolidLine(A.SchemeColorValues.Accent4, 0, 0);
                        }
                    }

                    PlotArea.UsedChartOptions[iChartType].HasHighLowLines = true;
                    if (IsStylish)
                    {
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Width = 0.75m;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CapType = A.LineCapValues.Flat;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CompoundLineType =
                            A.CompoundLineValues.Single;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Alignment = A.PenAlignmentValues.Center;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.SetSolidLine(A.SchemeColorValues.Text1,
                            0.25m, 0);
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.JoinType = SLLineJoinValues.Round;
                    }

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
                    PlotArea.HasSecondaryAxes = true;
                    PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;
                    PlotArea.SecondaryValueAxis.ForceAxisPosition = true;
                    PlotArea.SecondaryValueAxis.IsCrosses = true;
                    PlotArea.SecondaryValueAxis.Crosses = C.CrossesValues.Maximum;
                    PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
                    //this.PlotArea.SecondaryValueAxis.OtherAxisIsCrosses = true;
                    //this.PlotArea.SecondaryValueAxis.OtherAxisCrosses = C.CrossesValues.AutoZero;
                    PlotArea.SecondaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
                    //this.PlotArea.SecondaryTextAxis.IsCrosses = true;
                    //this.PlotArea.SecondaryTextAxis.Crosses = C.CrossesValues.AutoZero;
                    PlotArea.SecondaryTextAxis.OtherAxisIsCrosses = true;
                    PlotArea.SecondaryTextAxis.OtherAxisCrosses = C.CrossesValues.Maximum;

                    if (IsStylish)
                    {
                        PlotArea.SecondaryValueAxis.ShowMajorGridlines = true;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.Width = 0.75m;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.CapType = A.LineCapValues.Flat;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.CompoundLineType = A.CompoundLineValues.Single;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.Alignment = A.PenAlignmentValues.Center;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.JoinType = SLLineJoinValues.Round;

                        PlotArea.SecondaryTextAxis.ClearShapeProperties();
                    }

                    if (IsStylish) Legend.LegendPosition = C.LegendPositionValues.Bottom;
                    PlotArea.SetDataSeriesAutoAxisType();
                    ShowEmptyCellsAs = C.DisplayBlanksAsValues.Gap;
                    break;
                case SLStockChartType.VolumeOpenHighLowClose:
                    vType = SLDataSeriesChartType.StockChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                    iBarChartType = (int) SLDataSeriesChartType.BarChartColumnPrimary;
                    for (i = 0; i < PlotArea.DataSeries.Count; ++i)
                        if (i == 0)
                        {
                            PlotArea.DataSeries[i].ChartType = SLDataSeriesChartType.BarChartColumnPrimary;
                            if (IsStylish)
                            {
                                PlotArea.DataSeries[i].Options.Fill.SetSolidFill(A.SchemeColorValues.Accent1, 0, 0);
                                // 2.25 pt width
                                PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                                PlotArea.DataSeries[i].Options.Line.SetNoLine();
                            }

                            iBarChartType = (int) SLDataSeriesChartType.BarChartColumnPrimary;
                            PlotArea.UsedChartTypes[iBarChartType] = true;
                            PlotArea.UsedChartOptions[iBarChartType].BarDirection = C.BarDirectionValues.Column;
                            PlotArea.UsedChartOptions[iBarChartType].BarGrouping = C.BarGroupingValues.Clustered;
                            if (Options != null)
                            {
                                PlotArea.UsedChartOptions[iBarChartType].GapWidth = Options.GapWidth;
                                PlotArea.UsedChartOptions[iBarChartType].Overlap = Options.Overlap;
                                PlotArea.DataSeries[i].Options.ShapeProperties.Fill = Options.Fill.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.Outline = Options.Border.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.Shadow =
                                    Options.Shadow.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.Glow = Options.Glow.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.SoftEdge =
                                    Options.SoftEdge.Clone();
                                PlotArea.DataSeries[i].Options.ShapeProperties.Format3D = Options.Format3D.Clone();
                            }
                        }
                        else
                        {
                            PlotArea.DataSeries[i].ChartType = vType;
                            if (IsStylish)
                            {
                                // 2.25 pt width
                                PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                                PlotArea.DataSeries[i].Options.Line.CapType = A.LineCapValues.Round;
                                PlotArea.DataSeries[i].Options.Line.SetNoLine();
                                PlotArea.DataSeries[i].Options.Line.JoinType = SLLineJoinValues.Round;
                                PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.None;
                            }
                        }

                    PlotArea.UsedChartOptions[iChartType].HasHighLowLines = true;
                    PlotArea.UsedChartOptions[iChartType].HasUpDownBars = true;
                    if (IsStylish)
                    {
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Width = 0.75m;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CapType = A.LineCapValues.Flat;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CompoundLineType =
                            A.CompoundLineValues.Single;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Alignment = A.PenAlignmentValues.Center;
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.SetSolidLine(A.SchemeColorValues.Text1,
                            0.25m, 0);
                        PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.JoinType = SLLineJoinValues.Round;

                        PlotArea.UsedChartOptions[iChartType].UpDownBars.GapWidth = 150;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Fill.SetSolidFill(
                            A.SchemeColorValues.Light1, 0, 0);
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.Width = 0.75m;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.CapType = A.LineCapValues.Flat;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.CompoundLineType =
                            A.CompoundLineValues.Single;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.Alignment =
                            A.PenAlignmentValues.Center;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.SetSolidLine(
                            A.SchemeColorValues.Text1, 0.35m, 0);
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.JoinType =
                            SLLineJoinValues.Round;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Fill.SetSolidFill(
                            A.SchemeColorValues.Dark1, 0.25m, 0);
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.Width = 0.75m;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.CapType = A.LineCapValues.Flat;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.CompoundLineType =
                            A.CompoundLineValues.Single;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.Alignment =
                            A.PenAlignmentValues.Center;
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.SetSolidLine(
                            A.SchemeColorValues.Text1, 0.35m, 0);
                        PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.JoinType =
                            SLLineJoinValues.Round;
                    }

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
                    PlotArea.HasSecondaryAxes = true;
                    PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;
                    PlotArea.SecondaryValueAxis.ForceAxisPosition = true;
                    PlotArea.SecondaryValueAxis.IsCrosses = true;
                    PlotArea.SecondaryValueAxis.Crosses = C.CrossesValues.Maximum;
                    PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
                    //this.PlotArea.SecondaryValueAxis.OtherAxisIsCrosses = true;
                    //this.PlotArea.SecondaryValueAxis.OtherAxisCrosses = C.CrossesValues.AutoZero;
                    PlotArea.SecondaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
                    //this.PlotArea.SecondaryTextAxis.IsCrosses = true;
                    //this.PlotArea.SecondaryTextAxis.Crosses = C.CrossesValues.AutoZero;
                    PlotArea.SecondaryTextAxis.OtherAxisIsCrosses = true;
                    PlotArea.SecondaryTextAxis.OtherAxisCrosses = C.CrossesValues.Maximum;

                    if (IsStylish)
                    {
                        PlotArea.SecondaryValueAxis.ShowMajorGridlines = true;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.Width = 0.75m;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.CapType = A.LineCapValues.Flat;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.CompoundLineType = A.CompoundLineValues.Single;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.Alignment = A.PenAlignmentValues.Center;
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                        PlotArea.SecondaryValueAxis.MajorGridlines.Line.JoinType = SLLineJoinValues.Round;

                        PlotArea.SecondaryTextAxis.ClearShapeProperties();
                    }

                    if (IsStylish) Legend.LegendPosition = C.LegendPositionValues.Bottom;
                    PlotArea.SetDataSeriesAutoAxisType();
                    ShowEmptyCellsAs = C.DisplayBlanksAsValues.Gap;
                    break;
            }
        }

        /// <summary>
        ///     Set a doughnut chart using one of the built-in doughnut chart types.
        /// </summary>
        /// <param name="ChartType">A built-in doughnut chart type.</param>
        public void SetChartType(SLDoughnutChartType ChartType)
        {
            SetChartType(ChartType, null);
        }

        /// <summary>
        ///     Set a doughnut chart using one of the built-in doughnut chart types.
        /// </summary>
        /// <param name="ChartType">A built-in doughnut chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLDoughnutChartType ChartType, SLPieChartOptions Options)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLDoughnutChartType.Doughnut:
                    vType = SLDataSeriesChartType.DoughnutChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                    PlotArea.UsedChartOptions[iChartType].HoleSize = IsStylish ? (byte) 75 : (byte) 50;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);
                    break;
                case SLDoughnutChartType.ExplodedDoughnut:
                    vType = SLDataSeriesChartType.DoughnutChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                    PlotArea.UsedChartOptions[iChartType].HoleSize = IsStylish ? (byte) 75 : (byte) 50;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Explosion = 25;
                    break;
            }
        }

        /// <summary>
        ///     Set a scatter chart using one of the built-in scatter chart types.
        /// </summary>
        /// <param name="ChartType">A built-in scatter chart type.</param>
        public void SetChartType(SLScatterChartType ChartType)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLScatterChartType.ScatterWithOnlyMarkers:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                    {
                        ds.Options.Line.Width = 2.25m;
                        ds.Options.Line.SetNoLine();
                        if (IsStylish)
                        {
                            ds.Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            ds.Options.Marker.Size = 5;
                        }
                    }

                    SetPlotAreaValueAxes();
                    PlotArea.HasPrimaryAxes = true;

                    if (IsStylish)
                        PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    break;
                case SLScatterChartType.ScatterWithSmoothLinesAndMarkers:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.SmoothMarker;
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                    {
                        ds.Options.Smooth = true;
                        if (IsStylish)
                        {
                            ds.Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            ds.Options.Marker.Size = 5;
                        }
                    }

                    SetPlotAreaValueAxes();
                    PlotArea.HasPrimaryAxes = true;

                    if (IsStylish)
                        PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    break;
                case SLScatterChartType.ScatterWithSmoothLines:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.SmoothMarker;
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                    {
                        ds.Options.Smooth = true;
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;
                    }

                    SetPlotAreaValueAxes();
                    PlotArea.HasPrimaryAxes = true;

                    if (IsStylish)
                        PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    break;
                case SLScatterChartType.ScatterWithStraightLinesAndMarkers:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                    PlotArea.SetDataSeriesChartType(vType);

                    if (IsStylish)
                        for (var i = 0; i < PlotArea.DataSeries.Count; ++i)
                        {
                            PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }

                    SetPlotAreaValueAxes();
                    PlotArea.HasPrimaryAxes = true;

                    if (IsStylish)
                        PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    break;
                case SLScatterChartType.ScatterWithStraightLines:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;

                    SetPlotAreaValueAxes();
                    PlotArea.HasPrimaryAxes = true;

                    if (IsStylish)
                        PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    break;
            }
        }

        /// <summary>
        ///     Set an area chart using one of the built-in area chart types.
        /// </summary>
        /// <param name="ChartType">A built-in area chart type.</param>
        public void SetChartType(SLAreaChartType ChartType)
        {
            SetChartType(ChartType, null);
        }

        /// <summary>
        ///     Set an area chart using one of the built-in area chart types.
        /// </summary>
        /// <param name="ChartType">A built-in area chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLAreaChartType ChartType, SLAreaChartOptions Options)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLAreaChartType.Area:
                    vType = SLDataSeriesChartType.AreaChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    break;
                case SLAreaChartType.StackedArea:
                    vType = SLDataSeriesChartType.AreaChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    break;
                case SLAreaChartType.StackedAreaMax:
                    vType = SLDataSeriesChartType.AreaChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    PlotArea.PrimaryValueAxis.FormatCode = "0%";
                    break;
                case SLAreaChartType.Area3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Area3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    PlotArea.HasDepthAxis = true;
                    PlotArea.DepthAxis.IsCrosses = true;
                    PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
                case SLAreaChartType.StackedArea3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Area3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    break;
                case SLAreaChartType.StackedAreaMax3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Area3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    PlotArea.PrimaryValueAxis.FormatCode = "0%";
                    break;
            }
        }

        /// <summary>
        ///     Set a line chart using one of the built-in line chart types.
        /// </summary>
        /// <param name="ChartType">A built-in line chart type.</param>
        public void SetChartType(SLLineChartType ChartType)
        {
            SetChartType(ChartType, null);
        }

        /// <summary>
        ///     Set a line chart using one of the built-in line chart types.
        /// </summary>
        /// <param name="ChartType">A built-in line chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLLineChartType ChartType, SLLineChartOptions Options)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLLineChartType.Line:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.StackedLine:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.StackedLineMax:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    foreach (var ds in PlotArea.DataSeries)
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.LineWithMarkers:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    if (IsStylish)
                        for (var i = 0; i < PlotArea.DataSeries.Count; ++i)
                        {
                            PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.StackedLineWithMarkers:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    if (IsStylish)
                        for (var i = 0; i < PlotArea.DataSeries.Count; ++i)
                        {
                            PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.StackedLineWithMarkersMax:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    if (IsStylish)
                        for (var i = 0; i < PlotArea.DataSeries.Count; ++i)
                        {
                            PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.Line3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Line3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.HasDepthAxis = true;
                    PlotArea.DepthAxis.IsCrosses = true;
                    PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
            }
        }

        /// <summary>
        ///     Set a column chart using one of the built-in column chart types.
        /// </summary>
        /// <param name="ChartType">A built-in column chart type.</param>
        public void SetChartType(SLColumnChartType ChartType)
        {
            SetChartType(ChartType, null);
        }

        /// <summary>
        ///     Set a column chart using one of the built-in column chart types.
        /// </summary>
        /// <param name="ChartType">A built-in column chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLColumnChartType ChartType, SLBarChartOptions Options)
        {
            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLColumnChartType.ClusteredColumn:
                    vType = SLDataSeriesChartType.BarChartColumnPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;

                    if (IsStylish)
                    {
                        PlotArea.UsedChartOptions[iChartType].Overlap = -27;
                        PlotArea.UsedChartOptions[iChartType].GapWidth = 219;
                    }
                    break;
                case SLColumnChartType.StackedColumn:
                    vType = SLDataSeriesChartType.BarChartColumnPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedColumnMax:
                    vType = SLDataSeriesChartType.BarChartColumnPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.ClusteredColumn3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedColumn3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedColumnMax3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.Column3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Standard;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.HasDepthAxis = true;
                    break;
                case SLColumnChartType.ClusteredCylinder:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedCylinder:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedCylinderMax:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.Cylinder3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Standard;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.HasDepthAxis = true;
                    break;
                case SLColumnChartType.ClusteredCone:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedCone:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedConeMax:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.Cone3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Standard;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.HasDepthAxis = true;
                    break;
                case SLColumnChartType.ClusteredPyramid:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedPyramid:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedPyramidMax:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.Pyramid3D:
                    RotateX = 15;
                    RotateY = 20;
                    RightAngleAxes = false;
                    Perspective = 30;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Standard;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.HasDepthAxis = true;
                    break;
            }
        }

        /// <summary>
        ///     Set a bar chart using one of the built-in bar chart types.
        /// </summary>
        /// <param name="ChartType">A built-in bar chart type.</param>
        public void SetChartType(SLBarChartType ChartType)
        {
            SetChartType(ChartType, null);
        }

        /// <summary>
        ///     Set a bar chart using one of the built-in bar chart types.
        /// </summary>
        /// <param name="ChartType">A built-in bar chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLBarChartType ChartType, SLBarChartOptions Options)
        {
            // bar charts have their axis positions different from column charts.

            Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLBarChartType.ClusteredBar:
                    vType = SLDataSeriesChartType.BarChartBarPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish) PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                    break;
                case SLBarChartType.StackedBar:
                    vType = SLDataSeriesChartType.BarChartBarPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish) PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                    break;
                case SLBarChartType.StackedBarMax:
                    vType = SLDataSeriesChartType.BarChartBarPrimary;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish) PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                    break;
                case SLBarChartType.ClusteredBar3D:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedBar3D:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedBarMax3D:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.ClusteredHorizontalCylinder:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalCylinder:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalCylinderMax:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.ClusteredHorizontalCone:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalCone:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalConeMax:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.ClusteredHorizontalPyramid:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalPyramid:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalPyramidMax:
                    RotateX = 15;
                    RotateY = 20;
                    if (IsStylish) DepthPercent = 100;
                    RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int) vType;
                    PlotArea.UsedChartTypes[iChartType] = true;
                    PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    PlotArea.SetDataSeriesChartType(vType);

                    PlotArea.HasPrimaryAxes = true;
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (IsStylish)
                    {
                        PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        Floor.ClearShapeProperties();
                        Floor.Fill.SetNoFill();
                        Floor.Border.SetNoLine();
                    }
                    break;
            }
        }

        internal void SetPlotAreaAxes()
        {
            PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
            PlotArea.PrimaryTextAxis.AxisId = SLConstants.PrimaryAxis1;
            PlotArea.PrimaryTextAxis.Orientation = C.OrientationValues.MinMax;
            PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
            PlotArea.PrimaryTextAxis.FormatCode = SLConstants.NumberFormatGeneral;
            PlotArea.PrimaryTextAxis.SourceLinked = true;
            PlotArea.PrimaryTextAxis.HasNumberingFormat = true;
            PlotArea.PrimaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            PlotArea.PrimaryTextAxis.CrossingAxis = SLConstants.PrimaryAxis2;
            PlotArea.PrimaryTextAxis.IsCrosses = true;
            PlotArea.PrimaryTextAxis.Crosses = C.CrossesValues.AutoZero;
            PlotArea.PrimaryTextAxis.LabelAlignment = C.LabelAlignmentValues.Center;
            PlotArea.PrimaryTextAxis.LabelOffset = 100;
            PlotArea.PrimaryTextAxis.OtherAxisIsCrosses = true;
            PlotArea.PrimaryTextAxis.OtherAxisCrosses = C.CrossesValues.AutoZero;

            PlotArea.PrimaryValueAxis.AxisId = SLConstants.PrimaryAxis2;
            PlotArea.PrimaryValueAxis.Orientation = C.OrientationValues.MinMax;
            PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
            PlotArea.PrimaryValueAxis.ShowMajorGridlines = true;
            PlotArea.PrimaryValueAxis.FormatCode = SLConstants.NumberFormatGeneral;
            PlotArea.PrimaryValueAxis.SourceLinked = true;
            PlotArea.PrimaryValueAxis.HasNumberingFormat = true;
            PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            PlotArea.PrimaryValueAxis.CrossingAxis = SLConstants.PrimaryAxis1;
            PlotArea.PrimaryValueAxis.IsCrosses = true;
            PlotArea.PrimaryValueAxis.Crosses = C.CrossesValues.AutoZero;
            PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            PlotArea.PrimaryValueAxis.OtherAxisIsCrosses = true;
            PlotArea.PrimaryValueAxis.OtherAxisCrosses = C.CrossesValues.AutoZero;

            PlotArea.DepthAxis.AxisId = SLConstants.PrimaryAxis3;
            PlotArea.DepthAxis.Orientation = C.OrientationValues.MinMax;
            PlotArea.DepthAxis.AxisPosition = C.AxisPositionValues.Bottom;
            PlotArea.DepthAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            PlotArea.DepthAxis.CrossingAxis = SLConstants.PrimaryAxis2;

            PlotArea.SecondaryValueAxis.AxisId = SLConstants.SecondaryAxis2;
            PlotArea.SecondaryValueAxis.Orientation = C.OrientationValues.MinMax;
            PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;
            PlotArea.SecondaryValueAxis.FormatCode = SLConstants.NumberFormatGeneral;
            PlotArea.SecondaryValueAxis.SourceLinked = true;
            PlotArea.SecondaryValueAxis.HasNumberingFormat = true;
            PlotArea.SecondaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            PlotArea.SecondaryValueAxis.CrossingAxis = SLConstants.SecondaryAxis1;
            PlotArea.SecondaryValueAxis.IsCrosses = true;
            PlotArea.SecondaryValueAxis.Crosses = C.CrossesValues.Maximum;
            PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;

            PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
            PlotArea.SecondaryTextAxis.AxisId = SLConstants.SecondaryAxis1;
            PlotArea.SecondaryTextAxis.Orientation = C.OrientationValues.MinMax;
            PlotArea.SecondaryTextAxis.Delete = true;
            PlotArea.SecondaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
            PlotArea.SecondaryTextAxis.FormatCode = SLConstants.NumberFormatGeneral;
            PlotArea.SecondaryTextAxis.SourceLinked = true;
            PlotArea.SecondaryTextAxis.HasNumberingFormat = true;
            PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            PlotArea.SecondaryTextAxis.CrossingAxis = SLConstants.SecondaryAxis2;
            PlotArea.SecondaryTextAxis.LabelAlignment = C.LabelAlignmentValues.Center;
            PlotArea.SecondaryTextAxis.LabelOffset = 100;
            PlotArea.SecondaryTextAxis.OtherAxisIsCrosses = true;
            PlotArea.SecondaryTextAxis.OtherAxisCrosses = C.CrossesValues.Maximum;

            if (IsStylish)
            {
                PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.None;
                PlotArea.SecondaryValueAxis.MajorTickMark = C.TickMarkValues.None;
            }
        }

        /// <summary>
        ///     This assumes SetPlotAreaAxes() is already called so fewer properties are set.
        /// </summary>
        private void SetPlotAreaValueAxes()
        {
            PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Value;
            PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;

            PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;

            // secondary axes are not set because they're dependent on what you set
            // for the chart type plotted for the data series
        }

        /// <summary>
        ///     Set the position of the chart relative to the top-left of the worksheet.
        /// </summary>
        /// <param name="Top">
        ///     Top position of the chart based on row index. For example, 0.5 means at the half-way point of the 1st
        ///     row, 2.5 means at the half-way point of the 3rd row.
        /// </param>
        /// <param name="Left">
        ///     Left position of the chart based on column index. For example, 0.5 means at the half-way point of
        ///     the 1st column, 2.5 means at the half-way point of the 3rd column.
        /// </param>
        /// <param name="Bottom">
        ///     Bottom position of the chart based on row index. For example, 5.5 means at the half-way point of
        ///     the 6th row, 7.5 means at the half-way point of the 8th row.
        /// </param>
        /// <param name="Right">
        ///     Right position of the chart based on column index. For example, 5.5 means at the half-way point of
        ///     the 6th column, 7.5 means at the half-way point of the 8th column.
        /// </param>
        public void SetChartPosition(double Top, double Left, double Bottom, double Right)
        {
            double fTop = 0, fLeft = 0, fBottom = 1, fRight = 1;
            if (Top < Bottom)
            {
                fTop = Top;
                fBottom = Bottom;
            }
            else
            {
                fTop = Bottom;
                fBottom = fTop;
            }

            if (Left < Right)
            {
                fLeft = Left;
                fRight = Right;
            }
            else
            {
                fLeft = Right;
                fRight = Left;
            }

            if (fTop < 0.0) fTop = 0.0;
            if (fLeft < 0.0) fLeft = 0.0;
            if (fBottom >= SLConstants.RowLimit) fBottom = SLConstants.RowLimit;
            if (fRight >= SLConstants.ColumnLimit) fRight = SLConstants.ColumnLimit;

            TopPosition = fTop;
            LeftPosition = fLeft;
            BottomPosition = fBottom;
            RightPosition = fRight;
        }

        /// <summary>
        ///     Show the chart title.
        /// </summary>
        /// <param name="Overlay">True if the title overlaps the plot area. False otherwise.</param>
        public void ShowChartTitle(bool Overlay)
        {
            HasTitle = true;
            Title.Overlay = Overlay;
        }

        /// <summary>
        ///     Hide the chart title.
        /// </summary>
        public void HideChartTitle()
        {
            HasTitle = false;
        }

        /// <summary>
        ///     Show the chart legend.
        /// </summary>
        /// <param name="Position">Position of the legend. Default is Right.</param>
        /// <param name="Overlay">True if the legend overlaps the plot area. False otherwise.</param>
        public void ShowChartLegend(C.LegendPositionValues Position, bool Overlay)
        {
            ShowLegend = true;
            Legend.LegendPosition = Position;
            Legend.Overlay = Overlay;
        }

        /// <summary>
        ///     Hide the chart legend.
        /// </summary>
        public void HideChartLegend()
        {
            ShowLegend = false;
        }

        /// <summary>
        ///     Get the options for a specific data series.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <returns>The data series options for the specific data series. If the index is out of bounds, a default is returned.</returns>
        public SLDataSeriesOptions GetDataSeriesOptions(int DataSeriesIndex)
        {
            var dso = new SLDataSeriesOptions(listThemeColors);

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count))
                return dso;
            dso = PlotArea.DataSeries[index].Options.Clone();
            return dso;
        }

        /// <summary>
        ///     Set the options for a given data series.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="Options">The data series options.</param>
        public void SetDataSeriesOptions(int DataSeriesIndex, SLDataSeriesOptions Options)
        {
            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            PlotArea.DataSeries[index].Options = Options.Clone();
        }

        /// <summary>
        ///     Show the primary text (category/date/value) axis. This has no effect if the chart has no primary axes.
        /// </summary>
        public void ShowPrimaryTextAxis()
        {
            if (PlotArea.HasPrimaryAxes)
            {
                PlotArea.PrimaryTextAxis.Delete = false;
                PlotArea.PrimaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        ///     Hide the primary text (category/date/value) axis. This has no effect if the chart has no primary axes.
        /// </summary>
        public void HidePrimaryTextAxis()
        {
            if (PlotArea.HasPrimaryAxes)
            {
                PlotArea.PrimaryTextAxis.Delete = true;
                PlotArea.PrimaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        ///     Show the primary value axis. This has no effect if the chart has no primary axes.
        /// </summary>
        public void ShowPrimaryValueAxis()
        {
            if (PlotArea.HasPrimaryAxes)
            {
                PlotArea.PrimaryValueAxis.Delete = false;
                PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        ///     Hide the primary value axis. This has no effect if the chart has no primary axes.
        /// </summary>
        public void HidePrimaryValueAxis()
        {
            if (PlotArea.HasPrimaryAxes)
            {
                PlotArea.PrimaryValueAxis.Delete = true;
                PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        ///     Show the depth axis. This has no effect if the chart has no depth axis (that is, not a true 3D chart).
        /// </summary>
        public void ShowDepthAxis()
        {
            if (PlotArea.HasDepthAxis)
            {
                PlotArea.DepthAxis.Delete = false;
                PlotArea.DepthAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        ///     Hide the depth axis. This has no effect if the chart has no depth axis (that is, not a true 3D chart).
        /// </summary>
        public void HideDepthAxis()
        {
            if (PlotArea.HasDepthAxis)
            {
                PlotArea.DepthAxis.Delete = true;
                PlotArea.DepthAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        ///     Show the secondary text (category/date/value) axis. This has no effect if the chart has no secondary axes.
        /// </summary>
        public void ShowSecondaryTextAxis()
        {
            if (PlotArea.HasSecondaryAxes)
            {
                if (!HasShownSecondaryTextAxis)
                {
                    HasShownSecondaryTextAxis = true;
                    PlotArea.SecondaryTextAxis.AxisPosition =
                        SLChartTool.GetOppositePosition(PlotArea.SecondaryTextAxis.AxisPosition);
                }

                PlotArea.SecondaryTextAxis.Delete = false;
                PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        ///     Hide the secondary text (category/date/value) axis. This has no effect if the chart has no secondary axes.
        /// </summary>
        public void HideSecondaryTextAxis()
        {
            if (PlotArea.HasSecondaryAxes)
            {
                PlotArea.SecondaryTextAxis.Delete = true;
                PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        ///     Show the secondary value axis. This has no effect if the chart has no secondary axes.
        /// </summary>
        public void ShowSecondaryValueAxis()
        {
            if (PlotArea.HasSecondaryAxes)
            {
                PlotArea.SecondaryValueAxis.Delete = false;
                PlotArea.SecondaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        ///     Hide the secondary value axis. This has no effect if the chart has no secondary axes.
        /// </summary>
        public void HideSecondaryValueAxis()
        {
            if (PlotArea.HasSecondaryAxes)
            {
                PlotArea.SecondaryValueAxis.Delete = true;
                PlotArea.SecondaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        ///     Creates an instance of SLAreaChartOptions with theme information.
        /// </summary>
        /// <returns>An SLAreaChartOptions object with theme information.</returns>
        public SLAreaChartOptions CreateAreaChartOptions()
        {
            var aco = new SLAreaChartOptions(listThemeColors, IsStylish);
            return aco;
        }

        /// <summary>
        ///     Creates an instance of SLLineChartOptions with theme information.
        /// </summary>
        /// <returns>An SLLineChartOptions object with theme information.</returns>
        public SLLineChartOptions CreateLineChartOptions()
        {
            var lco = new SLLineChartOptions(listThemeColors, IsStylish);
            return lco;
        }

        /// <summary>
        ///     Creates an instance of SLPieChartOptions with theme information.
        /// </summary>
        /// <returns>An SLPieChartOptions object with theme information.</returns>
        public SLPieChartOptions CreatePieChartOptions()
        {
            var pco = new SLPieChartOptions(listThemeColors);
            if (IsStylish)
            {
                pco.Line.Width = 0.75m;
                pco.Line.CapType = A.LineCapValues.Flat;
                pco.Line.CompoundLineType = A.CompoundLineValues.Single;
                pco.Line.Alignment = A.PenAlignmentValues.Center;
                pco.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.65m, 0);
                pco.Line.JoinType = SLLineJoinValues.Round;
            }
            return pco;
        }

        /// <summary>
        ///     Creates an instance of SLStockChartOptions with theme information.
        /// </summary>
        /// <returns>An SLStockChartOptions object with theme information.</returns>
        public SLStockChartOptions CreateStockChartOptions()
        {
            var sco = new SLStockChartOptions(listThemeColors, IsStylish);
            return sco;
        }

        /// <summary>
        ///     Plot a specific data series as a doughnut chart. WARNING: Only weak checks done on whether the resulting
        ///     combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="ChartType">A built-in doughnut chart type for this specific data series.</param>
        public void PlotDataSeriesAsDoughnutChart(int DataSeriesIndex, SLDoughnutChartType ChartType)
        {
            PlotDataSeriesAsDoughnutChart(DataSeriesIndex, ChartType, null);
        }

        /// <summary>
        ///     Plot a specific data series as a doughnut chart. WARNING: Only weak checks done on whether the resulting
        ///     combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     Index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2
        ///     for the 2nd data series and so on.
        /// </param>
        /// <param name="ChartType">A built-in doughnut chart type for this specific data series.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsDoughnutChart(int DataSeriesIndex, SLDoughnutChartType ChartType,
            SLPieChartOptions Options)
        {
            // the original chart is not combinable
            if (!IsCombinable) return;

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            var vType = SLDataSeriesChartType.DoughnutChart;
            var iChartType = (int) vType;

            if (PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                // don't have to do anything if no options passed in.
                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                PlotArea.UsedChartTypes[iChartType] = true;
                PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                PlotArea.UsedChartOptions[iChartType].HoleSize = 50;
                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }

            PlotArea.DataSeries[index].ChartType = vType;

            switch (ChartType)
            {
                case SLDoughnutChartType.Doughnut:
                    PlotArea.DataSeries[index].Options.iExplosion = null;
                    break;
                case SLDoughnutChartType.ExplodedDoughnut:
                    PlotArea.DataSeries[index].Options.Explosion = 25;
                    break;
            }
        }

        /// <summary>
        ///     Plot a specific data series as a bar-of-pie chart. WARNING: Only weak checks done on whether the resulting
        ///     combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        public void PlotDataSeriesAsBarOfPieChart(int DataSeriesIndex)
        {
            PlotDataSeriesAsOfPieChart(DataSeriesIndex, true, null);
        }

        /// <summary>
        ///     Plot a specific data series as a bar-of-pie chart. WARNING: Only weak checks done on whether the resulting
        ///     combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsBarOfPieChart(int DataSeriesIndex, SLPieChartOptions Options)
        {
            PlotDataSeriesAsOfPieChart(DataSeriesIndex, true, Options);
        }

        /// <summary>
        ///     Plot a specific data series as a pie-of-pie chart. WARNING: Only weak checks done on whether the resulting
        ///     combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        public void PlotDataSeriesAsPieOfPieChart(int DataSeriesIndex)
        {
            PlotDataSeriesAsOfPieChart(DataSeriesIndex, false, null);
        }

        /// <summary>
        ///     Plot a specific data series as a pie-of-pie chart. WARNING: Only weak checks done on whether the resulting
        ///     combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPieOfPieChart(int DataSeriesIndex, SLPieChartOptions Options)
        {
            PlotDataSeriesAsOfPieChart(DataSeriesIndex, false, Options);
        }

        private void PlotDataSeriesAsOfPieChart(int DataSeriesIndex, bool IsBarOfPie, SLPieChartOptions Options)
        {
            // the original chart is not combinable
            if (!IsCombinable) return;

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            var vType = IsBarOfPie ? SLDataSeriesChartType.OfPieChartBar : SLDataSeriesChartType.OfPieChartPie;
            var iChartType = (int) vType;

            if (PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                // don't have to do anything if no options passed in.
                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                PlotArea.UsedChartTypes[iChartType] = true;
                PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                PlotArea.UsedChartOptions[iChartType].GapWidth = 100;
                PlotArea.UsedChartOptions[iChartType].SecondPieSize = 75;
                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }

            PlotArea.DataSeries[index].ChartType = vType;
        }

        /// <summary>
        ///     Plot a specific data series as a pie chart. WARNING: Only weak checks done on whether the resulting combination
        ///     chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="IsExploded">True to explode this data series. False otherwise.</param>
        public void PlotDataSeriesAsPieChart(int DataSeriesIndex, bool IsExploded)
        {
            PlotDataSeriesAsPieChart(DataSeriesIndex, IsExploded, null);
        }

        /// <summary>
        ///     Plot a specific data series as a pie chart. WARNING: Only weak checks done on whether the resulting combination
        ///     chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="IsExploded">True to explode this data series. False otherwise.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPieChart(int DataSeriesIndex, bool IsExploded, SLPieChartOptions Options)
        {
            // the original chart is not combinable
            if (!IsCombinable) return;

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            var vType = SLDataSeriesChartType.PieChart;
            var iChartType = (int) vType;

            if (PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                // don't have to do anything if no options passed in.
                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                PlotArea.UsedChartTypes[iChartType] = true;
                PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }

            PlotArea.DataSeries[index].ChartType = vType;

            if (IsExploded) PlotArea.DataSeries[index].Options.iExplosion = null;
            else PlotArea.DataSeries[index].Options.Explosion = 25;
        }

        /// <summary>
        ///     Plot a specific data series as a radar chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="ChartType">A built-in radar chart type for this specific data series.</param>
        public void PlotDataSeriesAsPrimaryRadarChart(int DataSeriesIndex, SLRadarChartType ChartType)
        {
            PlotDataSeriesAsRadarChart(DataSeriesIndex, ChartType, true);
        }

        /// <summary>
        ///     Plot a specific data series as a radar chart on the secondary axes. If there are no primary axes, it will be
        ///     plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is
        ///     valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="ChartType">A built-in radar chart type for this specific data series.</param>
        public void PlotDataSeriesAsSecondaryRadarChart(int DataSeriesIndex, SLRadarChartType ChartType)
        {
            PlotDataSeriesAsRadarChart(DataSeriesIndex, ChartType, false);
        }

        private void PlotDataSeriesAsRadarChart(int DataSeriesIndex, SLRadarChartType ChartType, bool IsPrimary)
        {
            // the original chart is not combinable
            if (!IsCombinable) return;

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            var bIsPrimary = IsPrimary;
            if (!PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                PlotArea.HasPrimaryAxes = true;
                PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
                PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;

                PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }
            else if (!bIsPrimary && PlotArea.HasPrimaryAxes && !PlotArea.HasSecondaryAxes)
            {
                PlotArea.HasSecondaryAxes = true;
                PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
                PlotArea.SecondaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                PlotArea.SecondaryTextAxis.ShowMajorGridlines = true;

                PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }

            // secondary radar: cat axis is also bottom, and value axis is also left, like the primary axis.

            var vType = bIsPrimary ? SLDataSeriesChartType.RadarChartPrimary : SLDataSeriesChartType.RadarChartSecondary;
            var iChartType = (int) vType;

            if (PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.
            }
            else
            {
                PlotArea.UsedChartTypes[iChartType] = true;

                switch (ChartType)
                {
                    case SLRadarChartType.Radar:
                        PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Marker;
                        PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        PlotArea.DataSeries[index].ChartType = vType;
                        break;
                    case SLRadarChartType.RadarWithMarkers:
                        PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Marker;
                        PlotArea.DataSeries[index].Options.Marker.vSymbol = null;
                        PlotArea.DataSeries[index].ChartType = vType;
                        break;
                    case SLRadarChartType.FilledRadar:
                        PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Filled;
                        PlotArea.DataSeries[index].Options.Marker.vSymbol = null;
                        PlotArea.DataSeries[index].ChartType = vType;
                        break;
                }
            }
        }

        /// <summary>
        ///     Plot a specific data series as an area chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        public void PlotDataSeriesAsPrimaryAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            PlotDataSeriesAsAreaChart(DataSeriesIndex, DisplayType, null, true);
        }

        /// <summary>
        ///     Plot a specific data series as an area chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPrimaryAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            SLAreaChartOptions Options)
        {
            PlotDataSeriesAsAreaChart(DataSeriesIndex, DisplayType, Options, true);
        }

        /// <summary>
        ///     Plot a specific data series as an area chart on the secondary axes. If there are no primary axes, it will be
        ///     plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is
        ///     valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        public void PlotDataSeriesAsSecondaryAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            PlotDataSeriesAsAreaChart(DataSeriesIndex, DisplayType, null, false);
        }

        /// <summary>
        ///     Plot a specific data series as an area chart on the secondary axes. If there are no primary axes, it will be
        ///     plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is
        ///     valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsSecondaryAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            SLAreaChartOptions Options)
        {
            PlotDataSeriesAsAreaChart(DataSeriesIndex, DisplayType, Options, false);
        }

        private void PlotDataSeriesAsAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            SLAreaChartOptions Options, bool IsPrimary)
        {
            // the original chart is not combinable
            if (!IsCombinable) return;

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            var bIsPrimary = IsPrimary;
            if (!PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                PlotArea.HasPrimaryAxes = true;
                PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
                PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;

                PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
            }
            else if (!bIsPrimary && PlotArea.HasPrimaryAxes && !PlotArea.HasSecondaryAxes)
            {
                PlotArea.HasSecondaryAxes = true;
                PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
                PlotArea.SecondaryTextAxis.AxisPosition = HasShownSecondaryTextAxis
                    ? C.AxisPositionValues.Top
                    : C.AxisPositionValues.Bottom;
                PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;

                PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
            }

            var vType = bIsPrimary ? SLDataSeriesChartType.AreaChartPrimary : SLDataSeriesChartType.AreaChartSecondary;
            var iChartType = (int) vType;

            if (PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                PlotArea.UsedChartTypes[iChartType] = true;

                switch (DisplayType)
                {
                    case SLChartDataDisplayType.Normal:
                        PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                    case SLChartDataDisplayType.Stacked:
                        PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                    case SLChartDataDisplayType.StackedMax:
                        PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                        if (bIsPrimary) PlotArea.PrimaryValueAxis.FormatCode = "0%";
                        else PlotArea.SecondaryValueAxis.FormatCode = "0%";
                        break;
                }

                PlotArea.DataSeries[index].ChartType = vType;
            }
        }

        /// <summary>
        ///     Plot a specific data series as a column chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        public void PlotDataSeriesAsPrimaryColumnChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, null, true, false);
        }

        /// <summary>
        ///     Plot a specific data series as a column chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPrimaryColumnChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            SLBarChartOptions Options)
        {
            PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, Options, true, false);
        }

        /// <summary>
        ///     Plot a specific data series as a column chart on the secondary axes. If there are no primary axes, it will be
        ///     plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is
        ///     valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        public void PlotDataSeriesAsSecondaryColumnChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, null, false, false);
        }

        /// <summary>
        ///     Plot a specific data series as a column chart on the secondary axes. If there are no primary axes, it will be
        ///     plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is
        ///     valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsSecondaryColumnChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            SLBarChartOptions Options)
        {
            PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, Options, false, false);
        }

        /// <summary>
        ///     Plot a specific data series as a bar chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        public void PlotDataSeriesAsPrimaryBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, null, true, true);
        }

        /// <summary>
        ///     Plot a specific data series as a bar chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPrimaryBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            SLBarChartOptions Options)
        {
            PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, Options, true, true);
        }

        /// <summary>
        ///     Plot a specific data series as a bar chart on the secondary axes. If there are no primary axes, it will be plotted
        ///     on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid.
        ///     Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        public void PlotDataSeriesAsSecondaryBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, null, false, true);
        }

        /// <summary>
        ///     Plot a specific data series as a bar chart on the secondary axes. If there are no primary axes, it will be plotted
        ///     on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid.
        ///     Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsSecondaryBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            SLBarChartOptions Options)
        {
            PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, Options, false, true);
        }

        private void PlotDataSeriesAsBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            SLBarChartOptions Options, bool IsPrimary, bool IsBar)
        {
            // the original chart is not combinable
            if (!IsCombinable) return;

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            var bIsPrimary = IsPrimary;
            if (!PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                PlotArea.HasPrimaryAxes = true;
                PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
                if (IsBar)
                {
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;
                }
                else
                {
                    PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                }

                PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }
            else if (!bIsPrimary && PlotArea.HasPrimaryAxes && !PlotArea.HasSecondaryAxes)
            {
                PlotArea.HasSecondaryAxes = true;
                PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
                if (IsBar)
                {
                    PlotArea.SecondaryTextAxis.AxisPosition = HasShownSecondaryTextAxis
                        ? C.AxisPositionValues.Right
                        : C.AxisPositionValues.Left;
                    PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Top;
                }
                else
                {
                    PlotArea.SecondaryTextAxis.AxisPosition = HasShownSecondaryTextAxis
                        ? C.AxisPositionValues.Top
                        : C.AxisPositionValues.Bottom;
                    PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;
                }

                PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }

            var vType = SLDataSeriesChartType.BarChartBarPrimary;
            if (bIsPrimary)
            {
                if (IsBar) vType = SLDataSeriesChartType.BarChartBarPrimary;
                else vType = SLDataSeriesChartType.BarChartColumnPrimary;
            }
            else
            {
                if (IsBar) vType = SLDataSeriesChartType.BarChartBarSecondary;
                else vType = SLDataSeriesChartType.BarChartColumnSecondary;
            }

            var iChartType = (int) vType;

            if (PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                PlotArea.UsedChartTypes[iChartType] = true;

                switch (DisplayType)
                {
                    case SLChartDataDisplayType.Normal:
                        PlotArea.UsedChartOptions[iChartType].BarDirection = IsBar
                            ? C.BarDirectionValues.Bar
                            : C.BarDirectionValues.Column;
                        PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                    case SLChartDataDisplayType.Stacked:
                        PlotArea.UsedChartOptions[iChartType].BarDirection = IsBar
                            ? C.BarDirectionValues.Bar
                            : C.BarDirectionValues.Column;
                        PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                        PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                    case SLChartDataDisplayType.StackedMax:
                        PlotArea.UsedChartOptions[iChartType].BarDirection = IsBar
                            ? C.BarDirectionValues.Bar
                            : C.BarDirectionValues.Column;
                        PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                        PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                }

                PlotArea.DataSeries[index].ChartType = vType;
            }
        }

        /// <summary>
        ///     Plot a specific data series as a scatter chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="ChartType">A built-in scatter chart type for this specific data series.</param>
        public void PlotDataSeriesAsPrimaryScatterChart(int DataSeriesIndex, SLScatterChartType ChartType)
        {
            PlotDataSeriesAsScatterChart(DataSeriesIndex, ChartType, true);
        }

        /// <summary>
        ///     Plot a specific data series as a scatter chart on the secondary axes. If there are no primary axes, it will be
        ///     plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is
        ///     valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="ChartType">A built-in scatter chart type for this specific data series.</param>
        public void PlotDataSeriesAsSecondaryScatterChart(int DataSeriesIndex, SLScatterChartType ChartType)
        {
            PlotDataSeriesAsScatterChart(DataSeriesIndex, ChartType, false);
        }

        private void PlotDataSeriesAsScatterChart(int DataSeriesIndex, SLScatterChartType ChartType, bool IsPrimary)
        {
            // the original chart is not combinable
            if (!IsCombinable) return;

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            var bIsPrimary = IsPrimary;
            if (!PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                PlotArea.HasPrimaryAxes = true;
                PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Value;
                PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;

                PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
            }
            else if (!bIsPrimary && PlotArea.HasPrimaryAxes && !PlotArea.HasSecondaryAxes)
            {
                PlotArea.HasSecondaryAxes = true;
                PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Value;
                PlotArea.SecondaryTextAxis.AxisPosition = HasShownSecondaryTextAxis
                    ? C.AxisPositionValues.Top
                    : C.AxisPositionValues.Bottom;
                PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                PlotArea.SecondaryTextAxis.ShowMajorGridlines = true;

                PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
            }

            var vType = bIsPrimary
                ? SLDataSeriesChartType.ScatterChartPrimary
                : SLDataSeriesChartType.ScatterChartSecondary;
            var iChartType = (int) vType;

            if (PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.
            }
            else
            {
                PlotArea.UsedChartTypes[iChartType] = true;

                switch (ChartType)
                {
                    case SLScatterChartType.ScatterWithOnlyMarkers:
                        PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                        PlotArea.DataSeries[index].ChartType = vType;
                        PlotArea.DataSeries[index].Options.Line.Width = 2.25m;
                        PlotArea.DataSeries[index].Options.Line.SetNoLine();
                        break;
                    case SLScatterChartType.ScatterWithSmoothLinesAndMarkers:
                        PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.SmoothMarker;
                        PlotArea.DataSeries[index].ChartType = vType;
                        PlotArea.DataSeries[index].Options.Smooth = true;
                        break;
                    case SLScatterChartType.ScatterWithSmoothLines:
                        PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.SmoothMarker;
                        PlotArea.DataSeries[index].ChartType = vType;
                        PlotArea.DataSeries[index].Options.Smooth = true;
                        PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                    case SLScatterChartType.ScatterWithStraightLinesAndMarkers:
                        PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                        PlotArea.DataSeries[index].ChartType = vType;
                        break;
                    case SLScatterChartType.ScatterWithStraightLines:
                        PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                        PlotArea.DataSeries[index].ChartType = vType;
                        PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                }
            }
        }

        /// <summary>
        ///     Plot a specific data series as a line chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="WithMarkers">True to display markers. False otherwise.</param>
        public void PlotDataSeriesAsPrimaryLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            bool WithMarkers)
        {
            PlotDataSeriesAsLineChart(DataSeriesIndex, DisplayType, WithMarkers, null, true);
        }

        /// <summary>
        ///     Plot a specific data series as a line chart on the primary axes. WARNING: Only weak checks done on whether the
        ///     resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="WithMarkers">True to display markers. False otherwise.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPrimaryLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            bool WithMarkers, SLLineChartOptions Options)
        {
            PlotDataSeriesAsLineChart(DataSeriesIndex, DisplayType, WithMarkers, Options, true);
        }

        /// <summary>
        ///     Plot a specific data series as a line chart on the secondary axes. If there are no primary axes, it will be plotted
        ///     on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid.
        ///     Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="WithMarkers">True to display markers. False otherwise.</param>
        public void PlotDataSeriesAsSecondaryLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            bool WithMarkers)
        {
            PlotDataSeriesAsLineChart(DataSeriesIndex, DisplayType, WithMarkers, null, false);
        }

        /// <summary>
        ///     Plot a specific data series as a line chart on the secondary axes. If there are no primary axes, it will be plotted
        ///     on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid.
        ///     Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DisplayType">
        ///     Chart display type. This corresponds to the 3 typical types in most charts: normal (or
        ///     clustered), stacked and 100% stacked.
        /// </param>
        /// <param name="WithMarkers">True to display markers. False otherwise.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsSecondaryLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType,
            bool WithMarkers, SLLineChartOptions Options)
        {
            PlotDataSeriesAsLineChart(DataSeriesIndex, DisplayType, WithMarkers, Options, false);
        }

        private void PlotDataSeriesAsLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, bool WithMarkers,
            SLLineChartOptions Options, bool IsPrimary)
        {
            // the original chart is not combinable
            if (!IsCombinable) return;

            var index = DataSeriesIndex - 1;

            // out of bounds
            if ((index < 0) || (index >= PlotArea.DataSeries.Count)) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            var bIsPrimary = IsPrimary;
            if (!PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                PlotArea.HasPrimaryAxes = true;
                PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
                PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;

                PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }
            else if (!bIsPrimary && PlotArea.HasPrimaryAxes && !PlotArea.HasSecondaryAxes)
            {
                PlotArea.HasSecondaryAxes = true;
                PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
                PlotArea.SecondaryTextAxis.AxisPosition = HasShownSecondaryTextAxis
                    ? C.AxisPositionValues.Top
                    : C.AxisPositionValues.Bottom;
                PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;

                PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }

            var vType = bIsPrimary ? SLDataSeriesChartType.LineChartPrimary : SLDataSeriesChartType.LineChartSecondary;
            var iChartType = (int) vType;

            if (PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                PlotArea.UsedChartTypes[iChartType] = true;

                switch (DisplayType)
                {
                    case SLChartDataDisplayType.Normal:
                        PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                        PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                        if (!WithMarkers) PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                    case SLChartDataDisplayType.Stacked:
                        PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                        PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                        if (!WithMarkers) PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                    case SLChartDataDisplayType.StackedMax:
                        PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                        PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                        if (Options != null) PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                        if (!WithMarkers) PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                }

                PlotArea.DataSeries[index].ChartType = vType;
            }
        }

        /// <summary>
        ///     Creates an instance of SLGroupDataLabelOptions with theme information.
        /// </summary>
        /// <returns>An SLGroupDataLabelOptions with theme information.</returns>
        public SLGroupDataLabelOptions CreateGroupDataLabelOptions()
        {
            return new SLGroupDataLabelOptions(listThemeColors);
        }

        /// <summary>
        ///     Creates an instance of SLDataLabelOptions with theme information.
        /// </summary>
        /// <returns>An SLDataLabelOptions with theme information.</returns>
        public SLDataLabelOptions CreateDataLabelOptions()
        {
            return new SLDataLabelOptions(listThemeColors);
        }

        /// <summary>
        ///     Set data label options to a specific data series.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="Options">Data label customization options.</param>
        public void SetGroupDataLabelOptions(int DataSeriesIndex, SLGroupDataLabelOptions Options)
        {
            // why not just return if outside of range? Because I assume you counted wrongly.
            if (DataSeriesIndex < 1) DataSeriesIndex = 1;
            if (DataSeriesIndex > PlotArea.DataSeries.Count) DataSeriesIndex = PlotArea.DataSeries.Count;
            // to get it to 0-index
            --DataSeriesIndex;

            PlotArea.DataSeries[DataSeriesIndex].GroupDataLabelOptions = Options.Clone();
        }

        /// <summary>
        ///     Set data label options to all data series.
        /// </summary>
        /// <param name="Options">Data label customization options.</param>
        public void SetGroupDataLabelOptions(SLGroupDataLabelOptions Options)
        {
            foreach (var ser in PlotArea.DataSeries)
                ser.GroupDataLabelOptions = Options.Clone();
        }

        /// <summary>
        ///     Set data label options to a specific data point in a specific data series.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DataPointIndex">
        ///     The index of the data point. This is 1-based indexing, so it's 1 for the 1st data point, 2
        ///     for the 2nd data point and so on.
        /// </param>
        /// <param name="Options">Data label customization options.</param>
        public void SetDataLabelOptions(int DataSeriesIndex, int DataPointIndex, SLDataLabelOptions Options)
        {
            // why not just return if outside of range? Because I assume you counted wrongly.
            if (DataSeriesIndex < 1) DataSeriesIndex = 1;
            if (DataSeriesIndex > PlotArea.DataSeries.Count) DataSeriesIndex = PlotArea.DataSeries.Count;
            // to get it to 0-index
            --DataSeriesIndex;

            --DataPointIndex;
            if (DataPointIndex < 0) DataPointIndex = 0;
            PlotArea.DataSeries[DataSeriesIndex].DataLabelOptionsList[DataPointIndex] = Options.Clone();
        }

        /// <summary>
        ///     Creates an instance of SLDataPointOptions with theme information.
        /// </summary>
        /// <returns>An SLDataPointOptions with theme information.</returns>
        public SLDataPointOptions CreateDataPointOptions()
        {
            return new SLDataPointOptions(listThemeColors);
        }

        /// <summary>
        ///     Set data point options to a specific data point in a specific data series.
        /// </summary>
        /// <param name="DataSeriesIndex">
        ///     The index of the data series. This is 1-based indexing, so it's 1 for the 1st data
        ///     series, 2 for the 2nd data series and so on.
        /// </param>
        /// <param name="DataPointIndex">
        ///     The index of the data point. This is 1-based indexing, so it's 1 for the 1st data point, 2
        ///     for the 2nd data point and so on.
        /// </param>
        /// <param name="Options">Data point customization options.</param>
        public void SetDataPointOptions(int DataSeriesIndex, int DataPointIndex, SLDataPointOptions Options)
        {
            // why not just return if outside of range? Because I assume you counted wrongly.
            if (DataSeriesIndex < 1) DataSeriesIndex = 1;
            if (DataSeriesIndex > PlotArea.DataSeries.Count) DataSeriesIndex = PlotArea.DataSeries.Count;
            // to get it to 0-index
            --DataSeriesIndex;

            --DataPointIndex;
            if (DataPointIndex < 0) DataPointIndex = 0;
            PlotArea.DataSeries[DataSeriesIndex].DataPointOptionsList[DataPointIndex] = Options.Clone();
        }

        internal C.ChartSpace ToChartSpace(ref ChartPart chartp)
        {
            ImagePart imgp;

            var cs = new C.ChartSpace();
            cs.AddNamespaceDeclaration("c", SLConstants.NamespaceC);
            cs.AddNamespaceDeclaration("a", SLConstants.NamespaceA);
            cs.AddNamespaceDeclaration("r", SLConstants.NamespaceRelationships);

            cs.Date1904 = new C.Date1904 {Val = Date1904};

            cs.EditingLanguage = new C.EditingLanguage();
            cs.EditingLanguage.Val = CultureInfo.CurrentCulture.Name;

            cs.RoundedCorners = new C.RoundedCorners {Val = RoundedCorners};

            var altcontent = new AlternateContent();
            altcontent.AddNamespaceDeclaration("mc", SLConstants.NamespaceMc);

            var altcontentchoice = new AlternateContentChoice {Requires = "c14"};
            altcontentchoice.AddNamespaceDeclaration("c14", SLConstants.NamespaceC14);
            // why +100? I don't know... ask Microsoft. But there are 48 styles. Even with the
            // advanced "+100" version, it's 96 total. It's a byte, with 256 possibilities.
            // As of this writing, Excel 2013 is rumoured to dispense away with this chart styling.
            // So maybe all this is moot anyway...
            altcontentchoice.Append(new C14.Style {Val = (byte) (ChartStyle + 100)});
            altcontent.Append(altcontentchoice);

            var altcontentfallback = new AlternateContentFallback();
            altcontentfallback.Append(new C.Style {Val = (byte) ChartStyle});
            altcontent.Append(altcontentfallback);

            cs.Append(altcontent);

            var chart = new C.Chart();

            if (HasView3D)
            {
                chart.View3D = new C.View3D();
                if (RotateX != null) chart.View3D.RotateX = new C.RotateX {Val = RotateX.Value};
                if (HeightPercent != null) chart.View3D.HeightPercent = new C.HeightPercent {Val = HeightPercent.Value};
                if (RotateY != null) chart.View3D.RotateY = new C.RotateY {Val = RotateY.Value};
                if (DepthPercent != null) chart.View3D.DepthPercent = new C.DepthPercent {Val = DepthPercent};
                if (RightAngleAxes != null)
                    chart.View3D.RightAngleAxes = new C.RightAngleAxes {Val = RightAngleAxes.Value};
                if (Perspective != null) chart.View3D.Perspective = new C.Perspective {Val = Perspective.Value};
            }

            if (HasTitle)
            {
                if (Title.Fill.Type == SLFillType.BlipFill)
                {
                    imgp = chartp.AddImagePart(SLDrawingTool.GetImagePartType(Title.Fill.BlipFileName));
                    using (var fs = new FileStream(Title.Fill.BlipFileName, FileMode.Open))
                    {
                        imgp.FeedData(fs);
                    }
                    Title.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                }
                chart.Title = Title.ToTitle(IsStylish);
            }
            else
            {
                chart.AutoTitleDeleted = new C.AutoTitleDeleted {Val = true};
            }

            if (Is3D)
            {
                chart.Floor = new C.Floor();
                chart.Floor.Thickness = new C.Thickness {Val = Floor.Thickness};
                if (Floor.ShapeProperties.HasShapeProperties)
                {
                    if (Floor.Fill.Type == SLFillType.BlipFill)
                    {
                        imgp = chartp.AddImagePart(SLDrawingTool.GetImagePartType(Floor.Fill.BlipFileName));
                        using (var fs = new FileStream(Floor.Fill.BlipFileName, FileMode.Open))
                        {
                            imgp.FeedData(fs);
                        }
                        Floor.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                    }
                    chart.Floor.ShapeProperties = Floor.ShapeProperties.ToCShapeProperties();
                }

                chart.SideWall = new C.SideWall();
                chart.SideWall.Thickness = new C.Thickness {Val = SideWall.Thickness};
                if (SideWall.ShapeProperties.HasShapeProperties)
                {
                    if (SideWall.Fill.Type == SLFillType.BlipFill)
                    {
                        imgp = chartp.AddImagePart(SLDrawingTool.GetImagePartType(SideWall.Fill.BlipFileName));
                        using (var fs = new FileStream(SideWall.Fill.BlipFileName, FileMode.Open))
                        {
                            imgp.FeedData(fs);
                        }
                        SideWall.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                    }
                    chart.SideWall.ShapeProperties = SideWall.ShapeProperties.ToCShapeProperties(IsStylish);
                }

                chart.BackWall = new C.BackWall();
                chart.BackWall.Thickness = new C.Thickness {Val = BackWall.Thickness};
                if (BackWall.ShapeProperties.HasShapeProperties)
                {
                    if (BackWall.Fill.Type == SLFillType.BlipFill)
                    {
                        imgp = chartp.AddImagePart(SLDrawingTool.GetImagePartType(BackWall.Fill.BlipFileName));
                        using (var fs = new FileStream(BackWall.Fill.BlipFileName, FileMode.Open))
                        {
                            imgp.FeedData(fs);
                        }
                        BackWall.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                    }
                    chart.BackWall.ShapeProperties = BackWall.ShapeProperties.ToCShapeProperties(IsStylish);
                }
            }

            if (PlotArea.Fill.Type == SLFillType.BlipFill)
            {
                imgp = chartp.AddImagePart(SLDrawingTool.GetImagePartType(PlotArea.Fill.BlipFileName));
                using (var fs = new FileStream(PlotArea.Fill.BlipFileName, FileMode.Open))
                {
                    imgp.FeedData(fs);
                }
                PlotArea.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
            }
            chart.PlotArea = PlotArea.ToPlotArea(IsStylish);

            if (ShowLegend)
            {
                if (Legend.Fill.Type == SLFillType.BlipFill)
                {
                    imgp = chartp.AddImagePart(SLDrawingTool.GetImagePartType(Legend.Fill.BlipFileName));
                    using (var fs = new FileStream(Legend.Fill.BlipFileName, FileMode.Open))
                    {
                        imgp.FeedData(fs);
                    }
                    Legend.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                }
                chart.Legend = Legend.ToLegend(IsStylish);
            }

            chart.PlotVisibleOnly = new C.PlotVisibleOnly {Val = !ShowHiddenData};

            chart.DisplayBlanksAs = new C.DisplayBlanksAs {Val = ShowEmptyCellsAs};

            chart.ShowDataLabelsOverMaximum = new C.ShowDataLabelsOverMaximum {Val = ShowDataLabelsOverMaximum};

            cs.Append(chart);

            if (ShapeProperties.HasShapeProperties)
                cs.Append(ShapeProperties.ToCShapeProperties(IsStylish));

            return cs;
        }

        internal SLChart Clone()
        {
            var chart = new SLChart();
            chart.listThemeColors = new List<Color>();
            for (var i = 0; i < listThemeColors.Count; ++i)
                chart.listThemeColors.Add(listThemeColors[i]);

            chart.Date1904 = Date1904;
            chart.IsStylish = IsStylish;
            chart.RoundedCorners = RoundedCorners;
            chart.IsCombinable = IsCombinable;

            chart.TopPosition = TopPosition;
            chart.LeftPosition = LeftPosition;
            chart.BottomPosition = BottomPosition;
            chart.RightPosition = RightPosition;
            chart.WorksheetName = WorksheetName;
            chart.RowsAsDataSeries = RowsAsDataSeries;
            chart.ShowHiddenData = ShowHiddenData;
            chart.ShowDataLabelsOverMaximum = ShowDataLabelsOverMaximum;

            chart.StartRowIndex = StartRowIndex;
            chart.StartColumnIndex = StartColumnIndex;
            chart.EndRowIndex = EndRowIndex;
            chart.EndColumnIndex = EndColumnIndex;

            chart.ChartStyle = ChartStyle;
            chart.ShowEmptyCellsAs = ShowEmptyCellsAs;
            chart.RotateX = RotateX;
            chart.HeightPercent = HeightPercent;
            chart.RotateY = RotateY;
            chart.DepthPercent = DepthPercent;
            chart.RightAngleAxes = RightAngleAxes;
            chart.Perspective = Perspective;
            chart.ChartName = ChartName;

            chart.HasTitle = HasTitle;
            chart.Title = Title.Clone();

            chart.Is3D = Is3D;

            chart.Floor = Floor.Clone();
            chart.SideWall = SideWall.Clone();
            chart.BackWall = BackWall.Clone();
            chart.PlotArea = PlotArea.Clone();
            chart.HasShownSecondaryTextAxis = HasShownSecondaryTextAxis;
            chart.ShowLegend = ShowLegend;
            chart.Legend = Legend.Clone();
            chart.ShapeProperties = ShapeProperties.Clone();

            return chart;
        }
    }
}