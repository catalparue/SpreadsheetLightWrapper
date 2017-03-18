using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Chart customization options for pie, bar-of-pie, pie-of-pie and doughnut charts.
    /// </summary>
    public class SLPieChartOptions
    {
        internal byte byHoleSize;

        internal bool HasSplit;

        internal ushort iFirstSliceAngle;

        internal ushort iGapWidth;

        internal ushort iSecondPieSize;

        internal SLShapeProperties ShapeProperties;

        /// <summary>
        ///     Initializes an instance of SLPieChartOptions. It is recommended to use SLChart.CreatePieChartOptions().
        /// </summary>
        public SLPieChartOptions()
        {
            Initialize(new List<Color>());
        }

        internal SLPieChartOptions(List<Color> ThemeColors)
        {
            Initialize(ThemeColors);
        }

        /// <summary>
        ///     Each data point shall have a different color. The default is "true".
        /// </summary>
        public bool VaryColors { get; set; }

        /// <summary>
        ///     Angle of the first slice, ranging from 0 degrees to 360 degrees.
        /// </summary>
        public ushort FirstSliceAngle
        {
            get { return iFirstSliceAngle; }
            set
            {
                iFirstSliceAngle = value;
                if (iFirstSliceAngle > 360) iFirstSliceAngle = 360;
            }
        }

        /// <summary>
        ///     The size of the hole in a doughnut chart, ranging from 10% to 90% of the diameter of the doughnut chart. If the
        ///     doughnut chart is exploded, the diameter is taken to be that when it's not exploded.
        /// </summary>
        public byte HoleSize
        {
            get { return byHoleSize; }
            set
            {
                byHoleSize = value;
                if (byHoleSize < 10) byHoleSize = 10;
                if (byHoleSize > 90) byHoleSize = 90;
            }
        }

        /// <summary>
        ///     The gap width between the first pie and the second bar or pie chart, ranging from 0 to 500 (both inclusive). This
        ///     is for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public ushort GapWidth
        {
            get { return iGapWidth; }
            set
            {
                iGapWidth = value;
                if (iGapWidth > 500) iGapWidth = 500;
            }
        }

        internal C.SplitValues SplitType { get; set; }
        internal double SplitPosition { get; set; }
        internal List<int> SecondPiePoints { get; set; }

        /// <summary>
        ///     The size of the second bar or pie of the bar-of-pie or pie-of-pie chart as a percentage of the size of the first
        ///     pie. This ranges from 5% to 200% (both inclusive).
        /// </summary>
        public ushort SecondPieSize
        {
            get { return iSecondPieSize; }
            set
            {
                iSecondPieSize = value;
                if (iSecondPieSize < 5) iSecondPieSize = 5;
                if (iSecondPieSize > 200) iSecondPieSize = 200;
            }
        }

        /// <summary>
        ///     Line properties for the connecting line for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public SLLinePropertiesType Line
        {
            get { return ShapeProperties.Outline; }
        }

        /// <summary>
        ///     Shadow properties for the connecting line for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public SLShadowEffect Shadow
        {
            get { return ShapeProperties.EffectList.Shadow; }
        }

        /// <summary>
        ///     Glow properties for the connecting line for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public SLGlow Glow
        {
            get { return ShapeProperties.EffectList.Glow; }
        }

        /// <summary>
        ///     Soft edge properties for the connecting line for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public SLSoftEdge SoftEdge
        {
            get { return ShapeProperties.EffectList.SoftEdge; }
        }

        private void Initialize(List<Color> ThemeColors)
        {
            VaryColors = true;
            iFirstSliceAngle = 0;
            byHoleSize = 10;
            iGapWidth = 150;
            HasSplit = false;
            SplitType = C.SplitValues.Position;
            SplitPosition = 0;
            SecondPiePoints = new List<int>();
            iSecondPieSize = 75;
            ShapeProperties = new SLShapeProperties(ThemeColors);
        }

        /// <summary>
        ///     Split the data series by position where the second plot contains the last N values. This is only for bar-of-pie or
        ///     pie-of-pie charts.
        /// </summary>
        /// <param name="LastNValues">The last N values used in the second plot.</param>
        public void SplitSeriesByPosition(int LastNValues)
        {
            HasSplit = true;
            SplitType = C.SplitValues.Position;
            SplitPosition = LastNValues;
            SecondPiePoints.Clear();
        }

        /// <summary>
        ///     Split the data series by value where the second plot contains all values less than a maximum value. This is only
        ///     for bar-of-pie or pie-of-pie charts.
        /// </summary>
        /// <param name="MaxValue">The maximum value.</param>
        public void SplitSeriesByValue(double MaxValue)
        {
            HasSplit = true;
            SplitType = C.SplitValues.Value;
            SplitPosition = MaxValue;
            SecondPiePoints.Clear();
        }

        /// <summary>
        ///     Split the data series by percentage where the second plot contains all values less than a percentage of the sum.
        ///     This is only for bar-of-pie or pie-of-pie charts.
        /// </summary>
        /// <param name="MaxPercentage">The maximum percentage of the sum.</param>
        public void SplitSeriesByPercentage(double MaxPercentage)
        {
            HasSplit = true;
            SplitType = C.SplitValues.Percent;
            SplitPosition = MaxPercentage;
            SecondPiePoints.Clear();
        }

        /// <summary>
        ///     Split the data series by selecting data points for the second plot. This is only for bar-of-pie or pie-of-pie
        ///     charts.
        /// </summary>
        /// <param name="DataPointIndices">
        ///     The indices of the data points of the data series. The index is 1-based, so "1,3,4" sets
        ///     the 1st, 3rd and 4th data point in the second plot.
        /// </param>
        public void SplitSeriesByCustom(params int[] DataPointIndices)
        {
            HasSplit = true;
            SplitType = C.SplitValues.Custom;
            SplitPosition = 0;
            SecondPiePoints.Clear();
            foreach (var i in DataPointIndices)
                if (i > 0) SecondPiePoints.Add(i - 1);
            SecondPiePoints.Sort();
        }

        internal SLPieChartOptions Clone()
        {
            var pco = new SLPieChartOptions(ShapeProperties.listThemeColors);
            pco.VaryColors = VaryColors;
            pco.iFirstSliceAngle = iFirstSliceAngle;
            pco.byHoleSize = byHoleSize;
            pco.iGapWidth = iGapWidth;
            pco.HasSplit = HasSplit;
            pco.SplitType = SplitType;
            pco.SplitPosition = SplitPosition;

            pco.SecondPiePoints = new List<int>();
            for (var i = 0; i < SecondPiePoints.Count; ++i)
                pco.SecondPiePoints.Add(SecondPiePoints[i]);

            pco.iSecondPieSize = iSecondPieSize;

            pco.ShapeProperties = ShapeProperties.Clone();

            return pco;
        }
    }
}