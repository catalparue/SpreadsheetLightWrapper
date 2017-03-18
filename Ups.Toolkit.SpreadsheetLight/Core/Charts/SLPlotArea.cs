using System;
using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    internal enum SLDataSeriesChartType
    {
        // make sure to start from 0 because this is also going to be used in array indices
        DoughnutChart = 0,
        OfPieChartBar,
        OfPieChartPie,
        PieChart,
        RadarChartPrimary,
        RadarChartSecondary,
        AreaChartPrimary,
        AreaChartSecondary,
        BarChartColumnPrimary,
        BarChartColumnSecondary,
        BarChartBarPrimary,
        BarChartBarSecondary,
        ScatterChartPrimary,
        ScatterChartSecondary,
        LineChartPrimary,
        LineChartSecondary,
        // the following supposedly can't be used in combination charts
        Area3DChart,
        Bar3DChart,
        BubbleChart,
        Line3DChart,
        Pie3DChart,
        SurfaceChart,
        Surface3DChart,
        StockChart,
        // just for default purposes. Shouldn't affect memory or performance just because there's one more enumeration.
        None
    }

    /// <summary>
    ///     Encapsulates properties and methods for setting plot areas in charts.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.PlotArea class.
    /// </summary>
    public class SLPlotArea
    {
        internal List<SLDataSeries> DataSeries;
        internal SLChartOptions[] UsedChartOptions;

        internal bool[] UsedChartTypes;

        internal SLPlotArea(List<Color> ThemeColors, bool Date1904, bool IsStylish = false)
        {
            InternalChartType = SLInternalChartType.Bar;

            var NumberOfChartTypes = Enum.GetNames(typeof(SLDataSeriesChartType)).Length;
            UsedChartTypes = new bool[NumberOfChartTypes];
            UsedChartOptions = new SLChartOptions[NumberOfChartTypes];
            for (var i = 0; i < NumberOfChartTypes; ++i)
            {
                UsedChartTypes[i] = false;
                UsedChartOptions[i] = new SLChartOptions(ThemeColors);
            }
            DataSeries = new List<SLDataSeries>();

            Layout = new SLLayout();

            PrimaryTextAxis = new SLTextAxis(ThemeColors, Date1904, IsStylish);
            PrimaryValueAxis = new SLValueAxis(ThemeColors, IsStylish);
            DepthAxis = new SLSeriesAxis(ThemeColors, IsStylish);
            SecondaryTextAxis = new SLTextAxis(ThemeColors, Date1904, IsStylish);
            SecondaryValueAxis = new SLValueAxis(ThemeColors, IsStylish);

            HasPrimaryAxes = false;
            HasDepthAxis = false;
            HasSecondaryAxes = false;

            ShowDataTable = false;
            DataTable = new SLDataTable(ThemeColors, IsStylish);

            ShapeProperties = new SLShapeProperties(ThemeColors);
            if (IsStylish)
            {
                ShapeProperties.Fill.SetNoFill();
                ShapeProperties.Outline.SetNoLine();
            }
        }

        internal SLInternalChartType InternalChartType { get; set; }

        internal SLLayout Layout { get; set; }

        internal SLTextAxis PrimaryTextAxis { get; set; }
        internal SLValueAxis PrimaryValueAxis { get; set; }
        internal SLSeriesAxis DepthAxis { get; set; }
        internal SLTextAxis SecondaryTextAxis { get; set; }
        internal SLValueAxis SecondaryValueAxis { get; set; }

        internal bool HasPrimaryAxes { get; set; }
        internal bool HasDepthAxis { get; set; }
        internal bool HasSecondaryAxes { get; set; }

        internal bool ShowDataTable { get; set; }
        internal SLDataTable DataTable { get; set; }

        internal SLShapeProperties ShapeProperties { get; set; }

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
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        internal void SetDataSeriesChartType(SLDataSeriesChartType ChartType)
        {
            for (var i = 0; i < DataSeries.Count; ++i)
                DataSeries[i].ChartType = ChartType;
        }

        internal void SetDataSeriesAutoAxisType()
        {
            // the first data series is good enough. In fact, AxisData should be identical for all.
            if (DataSeries.Count > 0)
                if (DataSeries[0].AxisData.UseNumberReference)
                {
                    var sFormatCode = DataSeries[0].AxisData.NumberReference.NumberingCache.FormatCode;
                    if (SLTool.CheckIfFormatCodeIsDateRelated(sFormatCode))
                    {
                        PrimaryTextAxis.AxisType = SLAxisType.Date;
                        PrimaryTextAxis.FormatCode = sFormatCode;
                        PrimaryTextAxis.BaseUnit = C.TimeUnitValues.Days;
                        SecondaryTextAxis.AxisType = SLAxisType.Date;
                        SecondaryTextAxis.FormatCode = sFormatCode;
                        SecondaryTextAxis.BaseUnit = C.TimeUnitValues.Days;
                    }
                }
        }

        internal C.PlotArea ToPlotArea(bool IsStylish = false)
        {
            var pa = new C.PlotArea();
            pa.Append(Layout.ToLayout());

            int iChartType;
            int i;

            // TODO: the rendering order is sort of listed in the following.
            // But apparently if you plot data series for doughnut first before bar-of-pie
            // it's different than if you plot bar-of-pie then doughnut.
            // Find out the "correct" order next version I suppose...

            // Excel 2010 apparently sets this by default for any chart...
            var gdlo = new SLGroupDataLabelOptions(ShapeProperties.listThemeColors);
            gdlo.ShowLegendKey = false;
            gdlo.ShowValue = false;
            gdlo.ShowCategoryName = false;
            gdlo.ShowSeriesName = false;
            gdlo.ShowPercentage = false;
            gdlo.ShowBubbleSize = false;

            #region Doughnut

            iChartType = (int) SLDataSeriesChartType.DoughnutChart;
            if (UsedChartTypes[iChartType])
            {
                var dc = new C.DoughnutChart();
                dc.VaryColors = new C.VaryColors {Val = UsedChartOptions[iChartType].VaryColors ?? true};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        dc.Append(DataSeries[i].ToPieChartSeries(IsStylish));

                dc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                dc.Append(new C.FirstSliceAngle {Val = UsedChartOptions[iChartType].FirstSliceAngle});
                dc.Append(new C.HoleSize {Val = UsedChartOptions[iChartType].HoleSize});

                pa.Append(dc);
            }

            #endregion

            #region Bar-of-pie

            iChartType = (int) SLDataSeriesChartType.OfPieChartBar;
            if (UsedChartTypes[iChartType])
            {
                var opc = new C.OfPieChart();
                opc.OfPieType = new C.OfPieType {Val = C.OfPieValues.Bar};
                opc.VaryColors = new C.VaryColors {Val = UsedChartOptions[iChartType].VaryColors ?? true};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        opc.Append(DataSeries[i].ToPieChartSeries(IsStylish));

                opc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                opc.Append(new C.GapWidth {Val = UsedChartOptions[iChartType].GapWidth});

                if (UsedChartOptions[iChartType].HasSplit)
                {
                    opc.Append(new C.SplitType {Val = UsedChartOptions[iChartType].SplitType});
                    if (UsedChartOptions[iChartType].SplitType != C.SplitValues.Custom)
                    {
                        opc.Append(new C.SplitPosition {Val = UsedChartOptions[iChartType].SplitPosition});
                    }
                    else
                    {
                        var custsplit = new C.CustomSplit();
                        foreach (var iPiePoint in UsedChartOptions[iChartType].SecondPiePoints)
                            custsplit.Append(new C.SecondPiePoint {Val = (uint) iPiePoint});
                        opc.Append(custsplit);
                    }
                }

                opc.Append(new C.SecondPieSize {Val = UsedChartOptions[iChartType].SecondPieSize});

                if (UsedChartOptions[iChartType].SeriesLinesShapeProperties.HasShapeProperties)
                    opc.Append(new C.SeriesLines
                    {
                        ChartShapeProperties =
                            UsedChartOptions[iChartType].SeriesLinesShapeProperties.ToChartShapeProperties()
                    });
                else
                    opc.Append(new C.SeriesLines());

                pa.Append(opc);
            }

            #endregion

            #region Pie-of-pie

            iChartType = (int) SLDataSeriesChartType.OfPieChartPie;
            if (UsedChartTypes[iChartType])
            {
                var opc = new C.OfPieChart();
                opc.OfPieType = new C.OfPieType {Val = C.OfPieValues.Pie};
                opc.VaryColors = new C.VaryColors {Val = UsedChartOptions[iChartType].VaryColors ?? true};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        opc.Append(DataSeries[i].ToPieChartSeries(IsStylish));

                opc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                opc.Append(new C.GapWidth {Val = UsedChartOptions[iChartType].GapWidth});

                if (UsedChartOptions[iChartType].HasSplit)
                {
                    opc.Append(new C.SplitType {Val = UsedChartOptions[iChartType].SplitType});
                    if (UsedChartOptions[iChartType].SplitType != C.SplitValues.Custom)
                    {
                        opc.Append(new C.SplitPosition {Val = UsedChartOptions[iChartType].SplitPosition});
                    }
                    else
                    {
                        var custsplit = new C.CustomSplit();
                        foreach (var iPiePoint in UsedChartOptions[iChartType].SecondPiePoints)
                            custsplit.Append(new C.SecondPiePoint {Val = (uint) iPiePoint});
                        opc.Append(custsplit);
                    }
                }

                opc.Append(new C.SecondPieSize {Val = UsedChartOptions[iChartType].SecondPieSize});

                if (UsedChartOptions[iChartType].SeriesLinesShapeProperties.HasShapeProperties)
                    opc.Append(new C.SeriesLines
                    {
                        ChartShapeProperties =
                            UsedChartOptions[iChartType].SeriesLinesShapeProperties.ToChartShapeProperties()
                    });
                else
                    opc.Append(new C.SeriesLines());

                pa.Append(opc);
            }

            #endregion

            #region Pie

            iChartType = (int) SLDataSeriesChartType.PieChart;
            if (UsedChartTypes[iChartType])
            {
                var pc = new C.PieChart();
                pc.VaryColors = new C.VaryColors {Val = UsedChartOptions[iChartType].VaryColors ?? true};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        pc.Append(DataSeries[i].ToPieChartSeries(IsStylish));

                pc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                pc.Append(new C.FirstSliceAngle {Val = UsedChartOptions[iChartType].FirstSliceAngle});

                pa.Append(pc);
            }

            #endregion

            #region Radar primary

            iChartType = (int) SLDataSeriesChartType.RadarChartPrimary;
            if (UsedChartTypes[iChartType])
            {
                var rc = new C.RadarChart();
                rc.RadarStyle = new C.RadarStyle {Val = UsedChartOptions[iChartType].RadarStyle};
                rc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        rc.Append(DataSeries[i].ToRadarChartSeries(IsStylish));

                rc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                rc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                rc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});

                pa.Append(rc);
            }

            #endregion

            #region Radar secondary

            iChartType = (int) SLDataSeriesChartType.RadarChartSecondary;
            if (UsedChartTypes[iChartType])
            {
                var rc = new C.RadarChart();
                rc.RadarStyle = new C.RadarStyle {Val = UsedChartOptions[iChartType].RadarStyle};
                rc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        rc.Append(DataSeries[i].ToRadarChartSeries(IsStylish));

                rc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                rc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis1});
                rc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis2});

                pa.Append(rc);
            }

            #endregion

            #region Area primary

            iChartType = (int) SLDataSeriesChartType.AreaChartPrimary;
            if (UsedChartTypes[iChartType])
            {
                var ac = new C.AreaChart();
                ac.Grouping = new C.Grouping {Val = UsedChartOptions[iChartType].Grouping};
                ac.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        ac.Append(DataSeries[i].ToAreaChartSeries(IsStylish));

                ac.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (UsedChartOptions[iChartType].HasDropLines)
                    ac.Append(UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));

                ac.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                ac.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});

                pa.Append(ac);
            }

            #endregion

            #region Area secondary

            iChartType = (int) SLDataSeriesChartType.AreaChartSecondary;
            if (UsedChartTypes[iChartType])
            {
                var ac = new C.AreaChart();
                ac.Grouping = new C.Grouping {Val = UsedChartOptions[iChartType].Grouping};
                ac.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        ac.Append(DataSeries[i].ToAreaChartSeries(IsStylish));

                ac.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (UsedChartOptions[iChartType].HasDropLines)
                    ac.Append(UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));

                ac.Append(new C.AxisId {Val = SLConstants.SecondaryAxis1});
                ac.Append(new C.AxisId {Val = SLConstants.SecondaryAxis2});

                pa.Append(ac);
            }

            #endregion

            #region Column primary

            iChartType = (int) SLDataSeriesChartType.BarChartColumnPrimary;
            if (UsedChartTypes[iChartType])
            {
                var bc = new C.BarChart();
                bc.BarDirection = new C.BarDirection {Val = C.BarDirectionValues.Column};
                bc.BarGrouping = new C.BarGrouping {Val = UsedChartOptions[iChartType].BarGrouping};
                bc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        bc.Append(DataSeries[i].ToBarChartSeries(IsStylish));

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth {Val = UsedChartOptions[iChartType].GapWidth});

                if (UsedChartOptions[iChartType].Overlap != 0)
                    bc.Append(new C.Overlap {Val = UsedChartOptions[iChartType].Overlap});

                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});

                pa.Append(bc);
            }

            #endregion

            #region Column secondary

            iChartType = (int) SLDataSeriesChartType.BarChartColumnSecondary;
            if (UsedChartTypes[iChartType])
            {
                var bc = new C.BarChart();
                bc.BarDirection = new C.BarDirection {Val = C.BarDirectionValues.Column};
                bc.BarGrouping = new C.BarGrouping {Val = UsedChartOptions[iChartType].BarGrouping};
                bc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        bc.Append(DataSeries[i].ToBarChartSeries(IsStylish));

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth {Val = UsedChartOptions[iChartType].GapWidth});

                if (UsedChartOptions[iChartType].Overlap != 0)
                    bc.Append(new C.Overlap {Val = UsedChartOptions[iChartType].Overlap});

                bc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis1});
                bc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis2});

                pa.Append(bc);
            }

            #endregion

            #region Bar primary

            iChartType = (int) SLDataSeriesChartType.BarChartBarPrimary;
            if (UsedChartTypes[iChartType])
            {
                var bc = new C.BarChart();
                bc.BarDirection = new C.BarDirection {Val = C.BarDirectionValues.Bar};
                bc.BarGrouping = new C.BarGrouping {Val = UsedChartOptions[iChartType].BarGrouping};
                bc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        bc.Append(DataSeries[i].ToBarChartSeries(IsStylish));

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth {Val = UsedChartOptions[iChartType].GapWidth});

                if (UsedChartOptions[iChartType].Overlap != 0)
                    bc.Append(new C.Overlap {Val = UsedChartOptions[iChartType].Overlap});

                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});

                pa.Append(bc);
            }

            #endregion

            #region Bar secondary

            iChartType = (int) SLDataSeriesChartType.BarChartBarSecondary;
            if (UsedChartTypes[iChartType])
            {
                var bc = new C.BarChart();
                bc.BarDirection = new C.BarDirection {Val = C.BarDirectionValues.Bar};
                bc.BarGrouping = new C.BarGrouping {Val = UsedChartOptions[iChartType].BarGrouping};
                bc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        bc.Append(DataSeries[i].ToBarChartSeries(IsStylish));

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth {Val = UsedChartOptions[iChartType].GapWidth});

                if (UsedChartOptions[iChartType].Overlap != 0)
                    bc.Append(new C.Overlap {Val = UsedChartOptions[iChartType].Overlap});

                bc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis1});
                bc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis2});

                pa.Append(bc);
            }

            #endregion

            #region Scatter primary

            iChartType = (int) SLDataSeriesChartType.ScatterChartPrimary;
            if (UsedChartTypes[iChartType])
            {
                var sc = new C.ScatterChart();
                sc.ScatterStyle = new C.ScatterStyle {Val = UsedChartOptions[iChartType].ScatterStyle};
                sc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        sc.Append(DataSeries[i].ToScatterChartSeries(IsStylish));

                sc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});

                pa.Append(sc);
            }

            #endregion

            #region Scatter secondary

            iChartType = (int) SLDataSeriesChartType.ScatterChartSecondary;
            if (UsedChartTypes[iChartType])
            {
                var sc = new C.ScatterChart();
                sc.ScatterStyle = new C.ScatterStyle {Val = UsedChartOptions[iChartType].ScatterStyle};
                sc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        sc.Append(DataSeries[i].ToScatterChartSeries(IsStylish));

                sc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                sc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis1});
                sc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis2});

                pa.Append(sc);
            }

            #endregion

            #region Line primary

            iChartType = (int) SLDataSeriesChartType.LineChartPrimary;
            if (UsedChartTypes[iChartType])
            {
                var lc = new C.LineChart();
                lc.Grouping = new C.Grouping {Val = UsedChartOptions[iChartType].Grouping};
                lc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        lc.Append(DataSeries[i].ToLineChartSeries(IsStylish));

                lc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (UsedChartOptions[iChartType].HasDropLines)
                    lc.Append(UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));

                lc.Append(new C.ShowMarker {Val = UsedChartOptions[iChartType].ShowMarker});
                lc.Append(new C.Smooth {Val = UsedChartOptions[iChartType].Smooth});

                lc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                lc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});

                pa.Append(lc);
            }

            #endregion

            #region Line secondary

            iChartType = (int) SLDataSeriesChartType.LineChartSecondary;
            if (UsedChartTypes[iChartType])
            {
                var lc = new C.LineChart();
                lc.Grouping = new C.Grouping {Val = UsedChartOptions[iChartType].Grouping};
                lc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        lc.Append(DataSeries[i].ToLineChartSeries(IsStylish));

                lc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (UsedChartOptions[iChartType].HasDropLines)
                    lc.Append(UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));

                lc.Append(new C.ShowMarker {Val = UsedChartOptions[iChartType].ShowMarker});
                lc.Append(new C.Smooth {Val = UsedChartOptions[iChartType].Smooth});

                lc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis1});
                lc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis2});

                pa.Append(lc);
            }

            #endregion

            #region Area3D

            iChartType = (int) SLDataSeriesChartType.Area3DChart;
            if (UsedChartTypes[iChartType])
            {
                var ac = new C.Area3DChart();
                ac.Grouping = new C.Grouping {Val = UsedChartOptions[iChartType].Grouping};
                ac.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        ac.Append(DataSeries[i].ToAreaChartSeries(IsStylish));

                ac.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (UsedChartOptions[iChartType].GapDepth != 150)
                    ac.Append(new C.GapDepth {Val = UsedChartOptions[iChartType].GapDepth});

                ac.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                ac.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});
                ac.Append(new C.AxisId {Val = SLConstants.PrimaryAxis3});

                pa.Append(ac);
            }

            #endregion

            #region Bar3D

            iChartType = (int) SLDataSeriesChartType.Bar3DChart;
            if (UsedChartTypes[iChartType])
            {
                var bc = new C.Bar3DChart();
                bc.BarDirection = new C.BarDirection {Val = UsedChartOptions[iChartType].BarDirection};
                bc.BarGrouping = new C.BarGrouping {Val = UsedChartOptions[iChartType].BarGrouping};
                bc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        bc.Append(DataSeries[i].ToBarChartSeries(IsStylish));

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth {Val = UsedChartOptions[iChartType].GapWidth});

                if (UsedChartOptions[iChartType].GapDepth != 150)
                    bc.Append(new C.GapDepth {Val = UsedChartOptions[iChartType].GapDepth});

                bc.Append(new C.Shape {Val = UsedChartOptions[iChartType].Shape});

                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});
                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis3});

                pa.Append(bc);
            }

            #endregion

            #region Bubble

            iChartType = (int) SLDataSeriesChartType.BubbleChart;
            if (UsedChartTypes[iChartType])
            {
                var bc = new C.BubbleChart();
                bc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        bc.Append(DataSeries[i].ToBubbleChartSeries(IsStylish));

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (!UsedChartOptions[iChartType].Bubble3D)
                    bc.Append(new C.Bubble3D {Val = UsedChartOptions[iChartType].Bubble3D});

                if (UsedChartOptions[iChartType].BubbleScale != 100)
                    bc.Append(new C.BubbleScale {Val = UsedChartOptions[iChartType].BubbleScale});

                if (!UsedChartOptions[iChartType].ShowNegativeBubbles)
                    bc.Append(new C.ShowNegativeBubbles {Val = UsedChartOptions[iChartType].ShowNegativeBubbles});

                if (UsedChartOptions[iChartType].SizeRepresents != C.SizeRepresentsValues.Area)
                    bc.Append(new C.SizeRepresents {Val = UsedChartOptions[iChartType].SizeRepresents});

                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                bc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});

                pa.Append(bc);
            }

            #endregion

            #region Line3D

            iChartType = (int) SLDataSeriesChartType.Line3DChart;
            if (UsedChartTypes[iChartType])
            {
                var lc = new C.Line3DChart();
                lc.Grouping = new C.Grouping {Val = UsedChartOptions[iChartType].Grouping};
                lc.VaryColors = new C.VaryColors {Val = false};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        lc.Append(DataSeries[i].ToLineChartSeries(IsStylish));

                lc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (UsedChartOptions[iChartType].HasDropLines)
                    lc.Append(UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));

                if (UsedChartOptions[iChartType].GapDepth != 150)
                    lc.Append(new C.GapDepth {Val = UsedChartOptions[iChartType].GapDepth});

                lc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                lc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});
                lc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis3});

                pa.Append(lc);
            }

            #endregion

            #region Pie3D

            iChartType = (int) SLDataSeriesChartType.Pie3DChart;
            if (UsedChartTypes[iChartType])
            {
                var pc = new C.Pie3DChart();
                pc.VaryColors = new C.VaryColors {Val = UsedChartOptions[iChartType].VaryColors ?? true};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        pc.Append(DataSeries[i].ToPieChartSeries(IsStylish));

                pc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                pa.Append(pc);
            }

            #endregion

            #region Surface

            iChartType = (int) SLDataSeriesChartType.SurfaceChart;
            if (UsedChartTypes[iChartType])
            {
                var sc = new C.SurfaceChart();
                if (UsedChartOptions[iChartType].bWireframe != null)
                    sc.Wireframe = new C.Wireframe {Val = UsedChartOptions[iChartType].Wireframe};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        sc.Append(DataSeries[i].ToSurfaceChartSeries(IsStylish));

                sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});
                sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis3});

                pa.Append(sc);
            }

            #endregion

            #region Surface3D

            iChartType = (int) SLDataSeriesChartType.Surface3DChart;
            if (UsedChartTypes[iChartType])
            {
                var sc = new C.Surface3DChart();
                if (UsedChartOptions[iChartType].bWireframe != null)
                    sc.Wireframe = new C.Wireframe {Val = UsedChartOptions[iChartType].Wireframe};

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        sc.Append(DataSeries[i].ToSurfaceChartSeries(IsStylish));

                sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});
                sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis3});

                pa.Append(sc);
            }

            #endregion

            #region Stock

            iChartType = (int) SLDataSeriesChartType.StockChart;
            if (UsedChartTypes[iChartType])
            {
                var sc = new C.StockChart();

                for (i = 0; i < DataSeries.Count; ++i)
                    if ((int) DataSeries[i].ChartType == iChartType)
                        sc.Append(DataSeries[i].ToLineChartSeries(IsStylish));

                sc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (UsedChartOptions[iChartType].HasDropLines)
                    sc.Append(UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));

                if (UsedChartOptions[iChartType].HasHighLowLines)
                    sc.Append(UsedChartOptions[iChartType].HighLowLines.ToHighLowLines(IsStylish));

                if (UsedChartOptions[iChartType].HasUpDownBars)
                    sc.Append(UsedChartOptions[iChartType].UpDownBars.ToUpDownBars(IsStylish));

                // stock charts either have a bar chart as the primary chart (the Volume) or doesn't.
                // If there is, then it's either a Volume-High-Low-Close or Volumn-Open-High-Low-Close,
                // so we use the secondary axis IDs.
                if (UsedChartTypes[(int) SLDataSeriesChartType.BarChartColumnPrimary])
                {
                    sc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis1});
                    sc.Append(new C.AxisId {Val = SLConstants.SecondaryAxis2});
                }
                else
                {
                    sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis1});
                    sc.Append(new C.AxisId {Val = SLConstants.PrimaryAxis2});
                }

                pa.Append(sc);
            }

            #endregion

            if (HasPrimaryAxes)
            {
                PrimaryTextAxis.IsCrosses = PrimaryValueAxis.OtherAxisIsCrosses;
                PrimaryTextAxis.Crosses = PrimaryValueAxis.OtherAxisCrosses;
                PrimaryTextAxis.CrossesAt = PrimaryValueAxis.OtherAxisCrossesAt;

                PrimaryTextAxis.OtherAxisIsInReverseOrder = PrimaryValueAxis.InReverseOrder;

                if ((PrimaryValueAxis.OtherAxisIsCrosses != null)
                    && PrimaryValueAxis.OtherAxisIsCrosses.Value
                    && (PrimaryValueAxis.OtherAxisCrosses == C.CrossesValues.Maximum))
                    PrimaryTextAxis.OtherAxisCrossedAtMaximum = true;
                else
                    PrimaryTextAxis.OtherAxisCrossedAtMaximum = false;

                PrimaryValueAxis.IsCrosses = PrimaryTextAxis.OtherAxisIsCrosses;
                PrimaryValueAxis.Crosses = PrimaryTextAxis.OtherAxisCrosses;
                PrimaryValueAxis.CrossesAt = PrimaryTextAxis.OtherAxisCrossesAt;

                PrimaryValueAxis.OtherAxisIsInReverseOrder = PrimaryTextAxis.InReverseOrder;

                if ((PrimaryTextAxis.OtherAxisIsCrosses != null)
                    && PrimaryTextAxis.OtherAxisIsCrosses.Value
                    && (PrimaryTextAxis.OtherAxisCrosses == C.CrossesValues.Maximum))
                    PrimaryValueAxis.OtherAxisCrossedAtMaximum = true;
                else
                    PrimaryValueAxis.OtherAxisCrossedAtMaximum = false;

                switch (PrimaryTextAxis.AxisType)
                {
                    case SLAxisType.Category:
                        pa.Append(PrimaryTextAxis.ToCategoryAxis(IsStylish));
                        break;
                    case SLAxisType.Date:
                        pa.Append(PrimaryTextAxis.ToDateAxis(IsStylish));
                        break;
                    case SLAxisType.Value:
                        pa.Append(PrimaryTextAxis.ToValueAxis(IsStylish));
                        break;
                }
                pa.Append(PrimaryValueAxis.ToValueAxis(IsStylish));
            }

            if (HasDepthAxis)
                pa.Append(DepthAxis.ToSeriesAxis(IsStylish));

            if (HasSecondaryAxes)
            {
                SecondaryTextAxis.IsCrosses = SecondaryValueAxis.OtherAxisIsCrosses;
                SecondaryTextAxis.Crosses = SecondaryValueAxis.OtherAxisCrosses;
                SecondaryTextAxis.CrossesAt = SecondaryValueAxis.OtherAxisCrossesAt;

                SecondaryTextAxis.OtherAxisIsInReverseOrder = SecondaryValueAxis.InReverseOrder;

                if ((SecondaryValueAxis.OtherAxisIsCrosses != null)
                    && SecondaryValueAxis.OtherAxisIsCrosses.Value
                    && (SecondaryValueAxis.OtherAxisCrosses == C.CrossesValues.Maximum))
                    SecondaryTextAxis.OtherAxisCrossedAtMaximum = true;
                else
                    SecondaryTextAxis.OtherAxisCrossedAtMaximum = false;

                SecondaryValueAxis.IsCrosses = SecondaryTextAxis.OtherAxisIsCrosses;
                SecondaryValueAxis.Crosses = SecondaryTextAxis.OtherAxisCrosses;
                SecondaryValueAxis.CrossesAt = SecondaryTextAxis.OtherAxisCrossesAt;

                SecondaryValueAxis.OtherAxisIsInReverseOrder = SecondaryTextAxis.InReverseOrder;

                if ((SecondaryTextAxis.OtherAxisIsCrosses != null)
                    && SecondaryTextAxis.OtherAxisIsCrosses.Value
                    && (SecondaryTextAxis.OtherAxisCrosses == C.CrossesValues.Maximum))
                    SecondaryValueAxis.OtherAxisCrossedAtMaximum = true;
                else
                    SecondaryValueAxis.OtherAxisCrossedAtMaximum = false;

                // the order of axes is:
                // 1) primary category/date/value axis
                // 2) primary value axis
                // 3) secondary value axis
                // 4) secondary category/date/value axis
                pa.Append(SecondaryValueAxis.ToValueAxis(IsStylish));
                switch (SecondaryTextAxis.AxisType)
                {
                    case SLAxisType.Category:
                        pa.Append(SecondaryTextAxis.ToCategoryAxis(IsStylish));
                        break;
                    case SLAxisType.Date:
                        pa.Append(SecondaryTextAxis.ToDateAxis(IsStylish));
                        break;
                    case SLAxisType.Value:
                        pa.Append(SecondaryTextAxis.ToValueAxis(IsStylish));
                        break;
                }
            }

            if (ShowDataTable) pa.Append(DataTable.ToDataTable(IsStylish));

            if (ShapeProperties.HasShapeProperties) pa.Append(ShapeProperties.ToChartShapeProperties(IsStylish));

            return pa;
        }

        internal SLPlotArea Clone()
        {
            var pa = new SLPlotArea(ShapeProperties.listThemeColors, PrimaryTextAxis.Date1904);
            pa.InternalChartType = InternalChartType;

            int i;

            pa.UsedChartTypes = new bool[UsedChartTypes.Length];
            for (i = 0; i < UsedChartTypes.Length; ++i)
                pa.UsedChartTypes[i] = UsedChartTypes[i];

            pa.UsedChartOptions = new SLChartOptions[UsedChartOptions.Length];
            for (i = 0; i < UsedChartOptions.Length; ++i)
                pa.UsedChartOptions[i] = UsedChartOptions[i].Clone();

            pa.DataSeries = new List<SLDataSeries>();
            for (i = 0; i < DataSeries.Count; ++i)
                pa.DataSeries.Add(DataSeries[i].Clone());

            pa.Layout = Layout.Clone();
            pa.PrimaryTextAxis = PrimaryTextAxis.Clone();
            pa.PrimaryValueAxis = PrimaryValueAxis.Clone();
            pa.DepthAxis = DepthAxis.Clone();
            pa.SecondaryTextAxis = SecondaryTextAxis.Clone();
            pa.SecondaryValueAxis = SecondaryValueAxis.Clone();
            pa.HasPrimaryAxes = HasPrimaryAxes;
            pa.HasDepthAxis = HasDepthAxis;
            pa.HasSecondaryAxes = HasSecondaryAxes;
            pa.ShowDataTable = ShowDataTable;
            pa.DataTable = DataTable.Clone();
            pa.ShapeProperties = ShapeProperties.Clone();

            return pa;
        }
    }
}