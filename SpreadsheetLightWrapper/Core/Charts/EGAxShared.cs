using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     For CategoryAxis, ValueAxis, SeriesAxis and DateAxis from namespace DocumentFormat.OpenXml.Drawing.Charts.
    /// </summary>
    public abstract class EGAxShared : SLChartAlignment
    {
        internal bool bSourceLinked;

        // Scaling
        internal double? fLogBase;

        // This is C.NumberingFormat
        internal bool HasNumberingFormat;

        internal bool? IsCrosses;
        internal bool OtherAxisCrossedAtMaximum;

        // The Excel UI sets cross values of that axis for the *other* axis.
        // Weird... Meaning if you set the cross value of the category axis (at least
        // on the UI), you're actually setting the cross value of the value axis.
        // Why Excel didn't set them on the actual axis is beyond me...
        // Maybe on the UI it made sense to do so.
        internal bool? OtherAxisIsCrosses;

        internal bool OtherAxisIsInReverseOrder;

        internal string sFormatCode;

        // C.ChartShapeProperties
        internal SLShapeProperties ShapeProperties;

        internal EGAxShared(List<Color> ThemeColors, bool IsStylish = false)
        {
            AxisId = 0;
            LogBase = null;
            Orientation = C.OrientationValues.MinMax;

            OtherAxisIsInReverseOrder = false;
            OtherAxisCrossedAtMaximum = false;

            MaxAxisValue = null;
            MinAxisValue = null;
            Delete = false;
            ForceAxisPosition = false;
            AxisPosition = C.AxisPositionValues.Bottom;

            ShowMajorGridlines = false;
            MajorGridlines = new SLMajorGridlines(ThemeColors, IsStylish);
            ShowMinorGridlines = false;
            MinorGridlines = new SLMinorGridlines(ThemeColors, IsStylish);

            ShowTitle = false;
            Title = new SLTitle(ThemeColors, IsStylish);

            sFormatCode = SLConstants.NumberFormatGeneral;
            bSourceLinked = true;
            HasNumberingFormat = false;

            MajorTickMark = C.TickMarkValues.Outside;
            MinorTickMark = C.TickMarkValues.None;
            TickLabelPosition = C.TickLabelPositionValues.NextTo; // default

            ShapeProperties = new SLShapeProperties(ThemeColors);

            CrossingAxis = 0;
            IsCrosses = null;
            Crosses = C.CrossesValues.AutoZero;
            CrossesAt = 0.0;

            OtherAxisIsCrosses = null;
            OtherAxisCrosses = C.CrossesValues.AutoZero;
            OtherAxisCrossesAt = 0.0;
        }

        internal uint AxisId { get; set; }

        internal double? LogBase
        {
            get { return fLogBase; }
            set
            {
                fLogBase = value;
                if (value != null)
                {
                    if (fLogBase < 2.0) fLogBase = 2.0;
                    if (fLogBase > 1000.0) fLogBase = 1000.0;
                }
            }
        }

        internal C.OrientationValues Orientation { get; set; }
        internal double? MaxAxisValue { get; set; }
        internal double? MinAxisValue { get; set; }

        /// <summary>
        ///     Display axis values in reverse order.
        /// </summary>
        public bool InReverseOrder
        {
            get { return Orientation == C.OrientationValues.MinMax ? false : true; }
            set
            {
                if (value) Orientation = C.OrientationValues.MaxMin;
                else Orientation = C.OrientationValues.MinMax;
            }
        }

        internal bool Delete { get; set; }

        internal bool ForceAxisPosition { get; set; }
        internal C.AxisPositionValues AxisPosition { get; set; }

        /// <summary>
        ///     Whether major gridlines are shown.
        /// </summary>
        public bool ShowMajorGridlines { get; set; }

        /// <summary>
        ///     Major gridlines properties.
        /// </summary>
        public SLMajorGridlines MajorGridlines { get; set; }

        /// <summary>
        ///     Whether minor gridlines are shown.
        /// </summary>
        public bool ShowMinorGridlines { get; set; }

        /// <summary>
        ///     Minor gridlines properties.
        /// </summary>
        public SLMinorGridlines MinorGridlines { get; set; }

        /// <summary>
        ///     Whether the axis title is shown.
        /// </summary>
        public bool ShowTitle { get; set; }

        /// <summary>
        ///     Axis title properties.
        /// </summary>
        public SLTitle Title { get; set; }

        /// <summary>
        ///     Format code for the axis. If you set a custom format code, you might also want to set SourceLinked to false.
        /// </summary>
        public string FormatCode
        {
            get { return sFormatCode; }
            set
            {
                sFormatCode = value;
                HasNumberingFormat = true;
            }
        }

        /// <summary>
        ///     Whether the format code is linked to the data source.
        /// </summary>
        public bool SourceLinked
        {
            get { return bSourceLinked; }
            set
            {
                bSourceLinked = value;
                HasNumberingFormat = true;
            }
        }

        /// <summary>
        ///     Major tick mark type.
        /// </summary>
        public C.TickMarkValues MajorTickMark { get; set; }

        /// <summary>
        ///     Minor tick mark type.
        /// </summary>
        public C.TickMarkValues MinorTickMark { get; set; }

        /// <summary>
        ///     Position of axis labels.
        /// </summary>
        public C.TickLabelPositionValues TickLabelPosition { get; set; }

        /// <summary>
        ///     Fill properties.
        /// </summary>
        public SLFill Fill
        {
            get { return ShapeProperties.Fill; }
        }

        /// <summary>
        ///     Line properties.
        /// </summary>
        public SLLinePropertiesType Line
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

        // C.TextProperties

        internal uint CrossingAxis { get; set; }
        internal C.CrossesValues Crosses { get; set; }
        internal double CrossesAt { get; set; }
        internal C.CrossesValues OtherAxisCrosses { get; set; }
        internal double OtherAxisCrossesAt { get; set; }
    }
}